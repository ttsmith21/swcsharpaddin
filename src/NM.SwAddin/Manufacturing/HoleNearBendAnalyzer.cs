using System;
using System.Collections.Generic;
using NM.Core;
using NM.Core.Manufacturing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using static NM.Core.Constants.UnitConversions;

namespace NM.SwAddin.Manufacturing
{
    /// <summary>
    /// Detects holes and cutouts that are too close to bend lines on the flat pattern.
    /// Operates on the flattened model: extracts bend line segments from the "Bend-Lines"
    /// sketch and inner loops (holes/cutouts) from the flat face, then computes the minimum
    /// 2D distance from each cutout edge to each bend line.
    ///
    /// Northern's rule: SS 304/316 = 5T + R, CRS/AL = 4T + R.
    /// For imported STEP files where bend radius can't be read, R defaults to 1.5T.
    /// </summary>
    public static class HoleNearBendAnalyzer
    {
        private const int SamplesPerEdge = 20;

        // Simple 2D point
        private struct Point2D
        {
            public double X, Y;
            public Point2D(double x, double y) { X = x; Y = y; }
        }

        // A bend line segment on the flat pattern
        private struct BendLineSegment
        {
            public Point2D Start, End;
            public double RadiusMeters; // per-bend radius (0 if unknown)
        }

        /// <summary>
        /// Analyze the flattened model for holes/cutouts too close to bend lines.
        /// The model must already be in the flat pattern state.
        /// </summary>
        /// <param name="model">The part model (must be flattened)</param>
        /// <param name="flatFace">The flat pattern face (largest planar face)</param>
        /// <param name="thicknessMeters">Sheet metal thickness in meters</param>
        /// <param name="material">Material string for lookup (e.g. "304 SS", "CRS")</param>
        /// <returns>Result with any violations found</returns>
        public static HoleNearBendResult Analyze(IModelDoc2 model, IFace2 flatFace,
            double thicknessMeters, string material)
        {
            var result = new HoleNearBendResult();
            if (model == null || flatFace == null || thicknessMeters <= 0)
                return result;

            try
            {
                ErrorHandler.DebugLog("[HoleNearBend] Starting analysis");

                // Step 1: Extract bend line segments with per-bend radii
                var bendLines = ExtractBendLines(model);
                if (bendLines.Count == 0)
                {
                    ErrorHandler.DebugLog("[HoleNearBend] No bend lines found");
                    return result;
                }
                ErrorHandler.DebugLog($"[HoleNearBend] Found {bendLines.Count} bend line(s)");

                // Step 2: Extract inner loops (holes/cutouts) from flat face
                var holeLoops = ExtractInnerLoops(flatFace);
                if (holeLoops.Count == 0)
                {
                    ErrorHandler.DebugLog("[HoleNearBend] No inner loops (holes/cutouts) found");
                    return result;
                }
                ErrorHandler.DebugLog($"[HoleNearBend] Found {holeLoops.Count} inner loop(s)");

                // Step 3: Check each hole against each bend line
                double thicknessIn = thicknessMeters * MetersToInches;
                double fallbackRadiusIn = 1.5 * thicknessIn; // R = 1.5T for imported STEP

                for (int h = 0; h < holeLoops.Count; h++)
                {
                    var holePoints = SampleLoopPoints(holeLoops[h]);
                    if (holePoints.Count == 0) continue;

                    for (int b = 0; b < bendLines.Count; b++)
                    {
                        var bend = bendLines[b];

                        // Per-bend radius: use actual if known, else fallback to 1.5T
                        double bendRadiusIn = bend.RadiusMeters > 0
                            ? bend.RadiusMeters * MetersToInches
                            : fallbackRadiusIn;

                        double requiredIn = BendClearanceLookup.ComputeMinDistance(
                            material, thicknessIn, bendRadiusIn);

                        // Find minimum distance from any hole edge point to this bend line
                        double minDist = double.MaxValue;
                        for (int p = 0; p < holePoints.Count; p++)
                        {
                            double d = PointToSegmentDistance2D(
                                holePoints[p],
                                bend.Start,
                                bend.End);
                            if (d < minDist) minDist = d;
                        }

                        // Convert to inches for comparison
                        double minDistIn = minDist * MetersToInches;

                        if (minDistIn < requiredIn)
                        {
                            int mult = BendClearanceLookup.GetMultiplier(material);
                            var violation = new HoleNearBendViolation
                            {
                                HoleIndex = h,
                                BendLineIndex = b,
                                ActualDistanceIn = Math.Round(minDistIn, 4),
                                RequiredDistanceIn = Math.Round(requiredIn, 4),
                                Description = string.Format(
                                    "Cutout #{0} is {1:F3}\" from bend line #{2}. " +
                                    "Minimum for {3} = {4:F3}\" ({5}T + R). " +
                                    "Feature will distort during bending. " +
                                    "Options: relocate feature >= {4:F3}\" from bend line, " +
                                    "or add relief slot between feature and bend.",
                                    h + 1, minDistIn, b + 1,
                                    string.IsNullOrWhiteSpace(material) ? "material" : material,
                                    requiredIn, mult)
                            };
                            result.Violations.Add(violation);
                            result.ViolationCount++;

                            ErrorHandler.DebugLog(string.Format(
                                "[HoleNearBend] VIOLATION: hole #{0} -> bend #{1}: " +
                                "{2:F4}\" actual vs {3:F4}\" required",
                                h + 1, b + 1, minDistIn, requiredIn));
                        }
                    }
                }

                ErrorHandler.DebugLog($"[HoleNearBend] Analysis complete: {result.ViolationCount} violation(s)");
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("HoleNearBendAnalyzer", ex.Message, ex, ErrorHandler.LogLevel.Warning);
            }

            return result;
        }

        /// <summary>
        /// Extract bend line segments from the "Bend-Lines" ProfileFeature sketch
        /// under the FlatPattern feature. Also collects per-bend radii from OneBend
        /// features under ProcessBends/FlattenBends containers.
        /// </summary>
        private static List<BendLineSegment> ExtractBendLines(IModelDoc2 model)
        {
            var lines = new List<BendLineSegment>();

            // Pass 1: Collect per-bend radii from OneBend features
            var bendRadii = CollectPerBendRadii(model);

            // Pass 2: Extract bend line geometry from the Bend-Lines sketch
            IFeature feat = model.FirstFeature() as IFeature;
            while (feat != null)
            {
                string type = feat.GetTypeName2() ?? string.Empty;
                if (type == "FlatPattern")
                {
                    var subFeat = feat.GetFirstSubFeature() as IFeature;
                    while (subFeat != null)
                    {
                        string subType = subFeat.GetTypeName2() ?? string.Empty;
                        string subName = subFeat.Name ?? string.Empty;

                        if (subType == "ProfileFeature" && subName.Length >= 10
                            && subName.Substring(0, 10) == "Bend-Lines")
                        {
                            ISketch swSketch = subFeat.GetSpecificFeature2() as ISketch;
                            if (swSketch != null)
                            {
                                object segObj = swSketch.GetSketchSegments();
                                if (segObj is object[] segs)
                                {
                                    for (int i = 0; i < segs.Length; i++)
                                    {
                                        var seg = segs[i] as ISketchSegment;
                                        if (seg == null) continue;

                                        try
                                        {
                                            if (seg.GetType() != (int)swSketchSegments_e.swSketchLINE)
                                                continue;

                                            var skLine = seg as ISketchLine;
                                            if (skLine == null) continue;

                                            var startPt = skLine.GetStartPoint2() as ISketchPoint;
                                            var endPt = skLine.GetEndPoint2() as ISketchPoint;
                                            if (startPt == null || endPt == null) continue;

                                            double radius = 0;
                                            if (i < bendRadii.Count)
                                                radius = bendRadii[i];

                                            lines.Add(new BendLineSegment
                                            {
                                                Start = new Point2D(startPt.X, startPt.Y),
                                                End = new Point2D(endPt.X, endPt.Y),
                                                RadiusMeters = radius
                                            });
                                        }
                                        catch { }
                                    }
                                }
                            }
                        }
                        subFeat = subFeat.GetNextSubFeature() as IFeature;
                    }
                }
                feat = feat.GetNextFeature() as IFeature;
            }

            return lines;
        }

        /// <summary>
        /// Collect bend radii from OneBend features in order.
        /// These correspond 1:1 with bend line sketch segments (same order).
        /// </summary>
        private static List<double> CollectPerBendRadii(IModelDoc2 model)
        {
            var radii = new List<double>();
            IFeature feat = model.FirstFeature() as IFeature;
            while (feat != null)
            {
                string type = feat.GetTypeName2() ?? string.Empty;
                if (type == "ProcessBends" || type == "FlattenBends")
                {
                    var subFeat = feat.GetFirstSubFeature() as IFeature;
                    while (subFeat != null)
                    {
                        if ((subFeat.GetTypeName2() ?? string.Empty) == "OneBend")
                        {
                            double r = 0;
                            try
                            {
                                var defObj = subFeat.GetDefinition();
                                if (defObj is IOneBendFeatureData one)
                                    r = one.BendRadius;
                            }
                            catch { }
                            radii.Add(r);
                        }
                        subFeat = subFeat.GetNextSubFeature() as IFeature;
                    }
                }
                feat = feat.GetNextFeature() as IFeature;
            }
            return radii;
        }

        /// <summary>
        /// Extract inner loops (holes, slots, cutouts) from the flat face.
        /// Inner loops are any loop where IsOuter() == false.
        /// </summary>
        private static List<ILoop2> ExtractInnerLoops(IFace2 flatFace)
        {
            var innerLoops = new List<ILoop2>();
            var loopsObj = flatFace.GetLoops() as object[];
            if (loopsObj == null) return innerLoops;

            for (int i = 0; i < loopsObj.Length; i++)
            {
                var loop = loopsObj[i] as ILoop2;
                if (loop == null) continue;
                if (!loop.IsOuter())
                    innerLoops.Add(loop);
            }
            return innerLoops;
        }

        /// <summary>
        /// Sample points along every edge in a loop.
        /// Returns points in model space (meters) for distance calculation.
        /// </summary>
        private static List<Point2D> SampleLoopPoints(ILoop2 loop)
        {
            var points = new List<Point2D>();
            var edges = loop.GetEdges() as object[];
            if (edges == null) return points;

            for (int e = 0; e < edges.Length; e++)
            {
                var edge = edges[e] as IEdge;
                if (edge == null) continue;

                var curve = edge.GetCurve() as ICurve;
                if (curve == null) continue;

                // Get parameter range for this edge
                var paramsObj = edge.GetCurveParams2() as object[];
                if (paramsObj == null || paramsObj.Length < 8) continue;

                double startParam = 0, endParam = 0;
                try
                {
                    double.TryParse(paramsObj[6]?.ToString(), out startParam);
                    double.TryParse(paramsObj[7]?.ToString(), out endParam);
                }
                catch { continue; }

                if (Math.Abs(endParam - startParam) < 1e-12) continue;

                // Sample evenly spaced points along the edge
                for (int s = 0; s <= SamplesPerEdge; s++)
                {
                    double t = startParam + (endParam - startParam) * s / SamplesPerEdge;
                    try
                    {
                        var evalResult = curve.Evaluate2(t, 0) as object[];
                        if (evalResult != null && evalResult.Length >= 3)
                        {
                            double x = Convert.ToDouble(evalResult[0]);
                            double y = Convert.ToDouble(evalResult[1]);
                            points.Add(new Point2D(x, y));
                        }
                    }
                    catch { }
                }
            }

            return points;
        }

        /// <summary>
        /// Compute the minimum distance from point P to line segment AB in 2D.
        /// All coordinates in the same unit (meters from SW API).
        /// </summary>
        private static double PointToSegmentDistance2D(Point2D p, Point2D a, Point2D b)
        {
            double abx = b.X - a.X;
            double aby = b.Y - a.Y;
            double apx = p.X - a.X;
            double apy = p.Y - a.Y;

            double dot = apx * abx + apy * aby;
            double lenSq = abx * abx + aby * aby;

            // Degenerate segment (zero length)
            if (lenSq < 1e-20)
            {
                return Math.Sqrt(apx * apx + apy * apy);
            }

            // Project P onto AB, clamped to [0,1]
            double t = dot / lenSq;
            if (t < 0) t = 0;
            else if (t > 1) t = 1;

            double closestX = a.X + t * abx;
            double closestY = a.Y + t * aby;

            double dx = p.X - closestX;
            double dy = p.Y - closestY;
            return Math.Sqrt(dx * dx + dy * dy);
        }
    }
}
