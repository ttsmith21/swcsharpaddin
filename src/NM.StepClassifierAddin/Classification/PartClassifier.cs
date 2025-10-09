using System;
using System.Diagnostics;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using NM.StepClassifierAddin.Utils;

namespace NM.StepClassifierAddin.Classification
{
    /// <summary>Implements the classification funnel.</summary>
    public static class PartClassifier
    {
        /// <summary>
        /// Classify a single solid body.
        /// </summary>
        public static PartPile Classify(ISldWorks app, IModelDoc2 model, IBody2 body)
        {
            if (app == null) throw new System.ArgumentNullException(nameof(app));
            if (model == null) throw new System.ArgumentNullException(nameof(model));
            if (body == null) throw new System.ArgumentNullException(nameof(body));

            var mp = model.Extension.CreateMassProperty() as IMassProperty;
            var m = GeometryScanner.Scan(body, mp);

            // Choose axis with highest side share
            int bestAxis = 0; double bestSide = -1;
            for (int i = 0; i < 3; i++)
            {
                double sideShare = m.ATotal > 0 ? m.Axes[i].SideArea / m.ATotal : 0.0;
                if (sideShare > bestSide) { bestSide = sideShare; bestAxis = i; }
            }
            double capShare = m.ATotal > 0 ? m.Axes[bestAxis].CapArea / m.ATotal : 0.0;
            double devShare = m.ATotal > 0 ? m.ADevelopable / m.ATotal : 0.0;

            // Complexity filter (reject very free-form)
            if (devShare < Thresholds.DEV_SURF_MIN_SHARE)
            {
                Log(model, m, bestAxis, bestSide, capShare, devShare, 0, 0, PartPile.Other);
                return PartPile.Other;
            }

            // SheetMetal via Utilities thickness; fallback to parallel planes when Utilities are missing
            double tMode, cov;
            bool hasUtilThickness = ThicknessAnalyzer.TryGetModeAndCoverage(app, body, out tMode, out cov) && tMode > 0;
            if (!hasUtilThickness)
            {
                hasUtilThickness = TryInferThicknessByParallelPlanes(body, out tMode, out cov) && tMode > 0;
            }
            // Fallback: infer thickness from coaxial cylindrical shell (bent sheet) when other methods fail
            if (!hasUtilThickness)
            {
                double shellThk;
                if (HasCoaxialCylindricalShell(body, out shellThk) && shellThk > 0)
                {
                    tMode = shellThk;
                    cov = 0.5; // conservative default; bend region usually not full coverage
                    hasUtilThickness = true;
                }
            }

            bool thinOk = (m.Spans.L2 > 0) && (tMode / Math.Max(m.Spans.L2, 1e-9) <= Thresholds.SHEET_THIN_RATIO_L2);
            double approxD = Math.Min(m.Spans.L2, m.Spans.L3); // crude diameter proxy for round-ish cross-sections
            double approxDIn = approxD * 39.3701;
            bool coverageOkPrimary = hasUtilThickness && ((cov >= Thresholds.SHEET_THICK_COVERAGE) || (devShare >= 0.95 && cov >= 0.10));
            bool coverageOkAny = hasUtilThickness && (coverageOkPrimary || (devShare >= 0.99 && cov >= 0.02));

            // Large-diameter override to SheetMetal before any Stick checks
            if (hasUtilThickness && thinOk && coverageOkPrimary && approxDIn >= 24.0)
            {
                Log(model, m, bestAxis, bestSide, capShare, devShare, tMode, cov, PartPile.SheetMetal);
                return PartPile.SheetMetal;
            }

            // Strong evidence for Stick/Tube: very high side share, very low caps, and constant section
            if (bestSide >= 0.97 && capShare <= 0.02)
            {
                bool constantSection = false;
                try { constantSection = Slicer.VerifyConstantSection(app, model, body, m.Axes[bestAxis].Axis, Thresholds.SECTION_AREA_TOL); } catch { }
                if (constantSection)
                {
                    Log(model, m, bestAxis, bestSide, capShare, devShare, 0, 0, PartPile.Stick);
                    return PartPile.Stick;
                }

                // Cylindrical shell fallback (e.g., pipe)
                double shellThk2;
                if (HasCoaxialCylindricalShell(body, out shellThk2))
                {
                    double ratio = approxD > 0 ? shellThk2 / (approxD * 0.5) : 0.0; // t / R
                    if (ratio >= 0.005 && ratio <= 0.25)
                    {
                        Log(model, m, bestAxis, bestSide, capShare, devShare, shellThk2, 1.0, PartPile.Stick);
                        return PartPile.Stick;
                    }
                }
            }

            // Default to SheetMetal when Utilities give a reasonable thin thickness
            if (hasUtilThickness && thinOk)
            {
                if (coverageOkPrimary || coverageOkAny)
                {
                    Log(model, m, bestAxis, bestSide, capShare, devShare, tMode, cov, PartPile.SheetMetal);
                    return PartPile.SheetMetal;
                }
            }

            Log(model, m, bestAxis, bestSide, capShare, devShare, tMode, cov, PartPile.Other);
            return PartPile.Other;
        }

        private static bool HasCoaxialCylindricalShell(IBody2 body, out double wallThk)
        {
            wallThk = 0.0;
            try
            {
                var faces = body.GetFaces() as object[]; if (faces == null || faces.Length == 0) return false;
                // Gather cylinders: axis (unit) and radius
                var cyls = new System.Collections.Generic.List<(double[] axis, double r)>();
                foreach (var fo in faces)
                {
                    var f = fo as IFace2; if (f == null) continue;
                    var s = f.IGetSurface(); if (s == null) continue;
                    bool isCyl = false; try { isCyl = s.IsCylinder(); } catch { }
                    if (!isCyl) continue;
                    try
                    {
                        var arrObj = s.CylinderParams as object; var arr = arrObj as double[];
                        if (arr != null && arr.Length >= 7)
                        {
                            var axis = Math3D.Normalize(new[] { arr[3], arr[4], arr[5] });
                            double r = Math.Abs(arr[6]);
                            if (r > 0) cyls.Add((axis, r));
                        }
                    }
                    catch { }
                }
                // Group by axis alignment and look for at least two different radii in a group
                const double AX_DOT_TOL = 0.999; // ~2.6°
                for (int i = 0; i < cyls.Count; i++)
                {
                    var ai = cyls[i].axis; double rmin = cyls[i].r, rmax = cyls[i].r; int count = 1;
                    for (int j = i + 1; j < cyls.Count; j++)
                    {
                        var aj = cyls[j].axis; double dot = Math.Abs(Math3D.Dot(ai, aj));
                        if (dot >= AX_DOT_TOL)
                        {
                            rmin = Math.Min(rmin, cyls[j].r); rmax = Math.Max(rmax, cyls[j].r); count++;
                        }
                    }
                    if (count >= 2)
                    {
                        double thk = rmax - rmin;
                        if (thk > 1e-5) { wallThk = thk; return true; }
                    }
                }
                return false;
            }
            catch { wallThk = 0.0; return false; }
        }

        private static bool TryInferThicknessByParallelPlanes(IBody2 body, out double thickness, out double coverage)
        {
            thickness = 0.0; coverage = 0.0;
            try
            {
                if (body == null) return false;
                var facesObj = body.GetFaces() as object[]; if (facesObj == null || facesObj.Length < 2) return false;

                // Gather planar faces with normals and a point on plane
                var items = new System.Collections.Generic.List<(IFace2 f, double area, double[] n, double[] p)>();
                double totalArea = 0.0;
                foreach (var o in facesObj)
                {
                    var f = o as IFace2; if (f == null) continue;
                    double a = 0.0; try { a = f.GetArea(); } catch { }
                    totalArea += a;
                    var s = f.IGetSurface(); if (s == null) continue;
                    bool isPlane = false; try { isPlane = s.IsPlane(); } catch { }
                    if (!isPlane) continue;
                    double[] n = null, p = null;
                    try
                    {
                        var arr = s.PlaneParams as object; var d = arr as double[];
                        if (d != null && d.Length >= 6)
                        { n = new[] { d[0], d[1], d[2] }; p = new[] { d[3], d[4], d[5] }; }
                    }
                    catch { }
                    if (n == null || p == null) continue;
                    items.Add((f, a, n, p));
                }
                if (items.Count < 2 || totalArea <= 0) return false;

                double bestScore = double.NegativeInfinity; double bestThk = 0.0; double bestArea = 0.0;
                for (int i = 0; i < items.Count; i++)
                {
                    for (int j = i + 1; j < items.Count; j++)
                    {
                        var a = items[i]; var b = items[j];
                        var an = Math3D.Normalize(a.n); var bn = Math3D.Normalize(b.n);
                        double dot = Math.Abs(Math3D.Dot(an, bn));
                        if (Math.Abs(dot - 1.0) > 0.05) continue; // nearly parallel planes
                        double[] d = { b.p[0] - a.p[0], b.p[1] - a.p[1], b.p[2] - a.p[2] };
                        double dist = Math.Abs(Math3D.Dot(d, an));
                        if (dist <= 0) continue;
                        double pairArea = Math.Min(a.area, b.area);
                        double score = pairArea - Math.Abs(1 - dot) * 1000.0;
                        if (score > bestScore)
                        {
                            bestScore = score; bestThk = dist; bestArea = pairArea;
                        }
                    }
                }
                if (bestThk > 0)
                {
                    thickness = bestThk;
                    coverage = Math.Max(0.0, Math.Min(1.0, bestArea / totalArea));
                    return true;
                }
                return false;
            }
            catch { thickness = 0.0; coverage = 0.0; return false; }
        }

        private static void Log(IModelDoc2 model, PrepassMetrics m, int bestAxis, double bestSide, double capShare, double devShare, double tMode, double cov, PartPile result)
        {
            try
            {
                string name = model?.GetTitle() ?? "(untitled)";
                NM.Core.ErrorHandler.DebugLog($"CLASSIFY '{name}': result={result}, side={bestSide:F3}, cap={capShare:F3}, dev={devShare:F3}, tMode={tMode*1000:F2}mm, cov={cov:F2}, L1={m.Spans.L1:F3}, L2={m.Spans.L2:F3}, L3={m.Spans.L3:F3}");
            }
            catch { }
        }
    }
}