using System;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using NM.Core;
using NM.SwAddin;

namespace NM.StepClassifierAddin.Classification
{
    /// <summary>
    /// Fast preflight checker for parts before attempting sheet-metal conversion.
    /// Pure analysis + labeling of problem parts; does not modify the model.
    /// </summary>
    public static class PartPreflight
    {
        // Heuristics/tolerances (meters)
        private const double KnifeEdgeLenTol = 1e-5;        // 0.01 mm
        private const double MicroFaceAreaTol = 1e-8;       // 0.01 mm^2
        private const double NonLinearEdgeRatioWarn = 0.50; // >50% of edges non-linear
        private const int NonLinearEdgeMinEdges = 50;       // only warn if enough edges scanned

        private static void DLog(string msg)
        {
            try { if (!string.IsNullOrWhiteSpace(msg)) ErrorHandler.DebugLog("Preflight: " + msg); } catch { }
        }

        public sealed class Result
        {
            public bool IsProblem { get; set; }
            public string Reason { get; set; }
            public int SolidBodyCount { get; set; }
            public int SurfaceBodyCount { get; set; }
            public int EdgeCount { get; set; }
            public int NonLinearEdgeCount { get; set; }
            public bool HasHelix { get; set; }
            public int KnifeEdgeCount { get; set; }
            public int MicroFaceCount { get; set; }
            public double Volume { get; set; }
            public bool HasSheetMetal { get; set; }
        }

        /// <summary>
        /// Analyze a model and optionally add an entry into the tracker when a problem is detected.
        /// Returns true when the part is OK for sheet-metal conversion.
        /// </summary>
        public static bool Evaluate(ModelInfo modelInfo, IModelDoc2 model, ProblemPartTracker tracker, out Result result)
        {
            result = Analyze(model);

            if (result.IsProblem && tracker != null && modelInfo != null)
            {
                string reason = BuildReasonText(result);
                tracker.AddProblemPart(modelInfo, reason);
                modelInfo.MarkProblem(reason);
                modelInfo.NeedsFix = true;
            }
            else if (!result.IsProblem && modelInfo != null)
            {
                modelInfo.NeedsFix = false;
            }

            return !result.IsProblem;
        }

        /// <summary>
        /// Perform preflight analysis without any side-effects.
        /// </summary>
        public static Result Analyze(IModelDoc2 model)
        {
            var res = new Result { IsProblem = false, Reason = string.Empty };
            const string proc = nameof(Analyze);
            ErrorHandler.PushCallStack(proc);
            try
            {
                string title = string.Empty; try { title = model?.GetTitle() ?? string.Empty; } catch { }
                DLog($"Start: title='{title}'");

                if (model == null)
                {
                    res.IsProblem = true; res.Reason = "Null model"; DLog("Fail: null model"); return res;
                }

                if (model.GetType() != (int)swDocumentTypes_e.swDocPART)
                {
                    res.IsProblem = true; res.Reason = "Not a part document"; DLog("Fail: not a part"); return res;
                }

                var part = model as IPartDoc;
                if (part == null)
                {
                    res.IsProblem = true; res.Reason = "Model cast to IPartDoc failed"; DLog("Fail: cast to IPartDoc failed"); return res;
                }

                // Body counts
                int solidCount = GetBodyCount(part, swBodyType_e.swSolidBody);
                int surfaceCount = GetBodyCount(part, swBodyType_e.swSheetBody);
                res.SolidBodyCount = solidCount;
                res.SurfaceBodyCount = surfaceCount;
                DLog($"Bodies: solids={solidCount}, surfaces={surfaceCount}");

                if (solidCount <= 0 && surfaceCount <= 0)
                { res.IsProblem = true; res.Reason = "Empty file: no bodies"; DLog("Fail: empty file (no bodies)"); return res; }
                if (solidCount <= 0 && surfaceCount > 0)
                { res.IsProblem = true; res.Reason = "Surface bodies only: no solid body"; DLog("Fail: surface-only"); return res; }
                if (solidCount > 1)
                { res.IsProblem = true; res.Reason = $"Multiple solid bodies ({solidCount})"; DLog(res.Reason); return res; }
                if (solidCount == 1 && surfaceCount > 0)
                { res.IsProblem = true; res.Reason = $"Mixed bodies: 1 solid + {surfaceCount} surface bodies"; DLog(res.Reason); return res; }

                // Zero/invalid volume check
                double vol = SolidWorksApiWrapper.GetModelVolume(model);
                res.Volume = vol;
                DLog($"Volume: {vol:E6} m^3");
                if (vol <= 0)
                { res.IsProblem = true; res.Reason = "Zero or invalid volume"; DLog("Fail: zero/invalid volume"); return res; }

                // Sheet metal feature presence (used to relax certain complexity flags)
                res.HasSheetMetal = SolidWorksApiWrapper.HasSheetMetalFeature(model);
                DLog($"HasSheetMetal={res.HasSheetMetal}");

                // Geometry complexity checks on the single solid body
                var bodies = part.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
                var body = bodies != null && bodies.Length > 0 ? bodies[0] as IBody2 : null;
                if (body == null)
                {
                    res.IsProblem = true; res.Reason = "Failed to get main solid body"; DLog("Fail: no main body"); return res;
                }

                AnalyzeEdgesUnique(body, res);
                AnalyzeFaces(body, res);
                DLog($"Edges: total={res.EdgeCount}, nonLinear={res.NonLinearEdgeCount}, knife={res.KnifeEdgeCount}, hasHelix={res.HasHelix}");
                DLog($"Faces: micro={res.MicroFaceCount}");

                // Hard-stops regardless of SM feature:
                if (res.HasHelix)
                {
                    res.IsProblem = true; res.Reason = "Contains helical/spiral curves (likely threads)"; DLog("Fail: helix/spiral"); return res;
                }

                // If already a sheet-metal part, skip soft-complexity flags
                if (!res.HasSheetMetal)
                {
                    if (res.EdgeCount >= NonLinearEdgeMinEdges)
                    {
                        double ratio = res.EdgeCount > 0 ? (double)res.NonLinearEdgeCount / res.EdgeCount : 0.0;
                        DLog($"NonLinear ratio={ratio:P2} (threshold={NonLinearEdgeRatioWarn:P0}, minEdges={NonLinearEdgeMinEdges})");
                        if (ratio >= NonLinearEdgeRatioWarn)
                        {
                            res.IsProblem = true; res.Reason = $"High non-linear edge ratio: {ratio:P0}"; DLog("Fail: " + res.Reason); return res;
                        }
                    }
                    if (res.MicroFaceCount >= 20)
                    {
                        res.IsProblem = true; res.Reason = $"Many micro faces (< {MicroFaceAreaTol * 1e6:0.###} mm^2): {res.MicroFaceCount}"; DLog("Fail: " + res.Reason); return res;
                    }
                }

                // Knife edges are suspicious even for SM, but keep them as soft: do not hard-fail unless there are many
                if (res.KnifeEdgeCount >= 50)
                {
                    res.IsProblem = true; res.Reason = $"Many tiny edges (< {KnifeEdgeLenTol * 1000:0.###} mm): {res.KnifeEdgeCount}"; DLog("Fail: " + res.Reason); return res;
                }

                // Passed all preflight checks
                res.IsProblem = false;
                res.Reason = string.Empty;
                DLog("Success: preflight OK");
                return res;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Preflight exception", ex);
                res.IsProblem = true;
                res.Reason = ex.Message;
                return res;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        private static int GetBodyCount(IPartDoc part, swBodyType_e type)
        {
            try
            {
                var arr = part.GetBodies2((int)type, true) as object[];
                return arr?.Length ?? 0;
            }
            catch { return 0; }
        }

        // Unique edge analysis across the body (avoids double-counting edges per face)
        private static void AnalyzeEdgesUnique(IBody2 body, Result res)
        {
            try
            {
                var edges = body.GetEdges() as object[]; if (edges == null) return;
                int total = 0, nonLinear = 0, knife = 0; bool hasHelix = false;

                foreach (var eo in edges)
                {
                    var e = eo as IEdge; if (e == null) continue;
                    total++;

                    double len = GetEdgeChordLength(e);
                    if (len > 0 && len <= KnifeEdgeLenTol) knife++;

                    var c = e.GetCurve() as ICurve;
                    bool isLine    = InvokeCurveBool(c, "IsLine",   defaultValue: true); // conservative default: treat unknown as simple
                    bool isCircle  = InvokeCurveBool(c, "IsCircle");
                    bool isArc     = InvokeCurveBool(c, "IsArc");
                    bool isEllipse = InvokeCurveBool(c, "IsEllipse");
                    bool helix     = InvokeCurveBool(c, "IsHelix");
                    if (helix) hasHelix = true;

                    bool simple = isLine || isCircle || isArc || isEllipse;
                    if (!simple) nonLinear++;
                }

                res.EdgeCount = total;
                res.NonLinearEdgeCount = nonLinear;
                res.KnifeEdgeCount = knife;
                res.HasHelix = hasHelix;
            }
            catch { }
        }

        private static void AnalyzeFaces(IBody2 body, Result res)
        {
            try
            {
                var faces = body.GetFaces() as object[]; if (faces == null) return;
                int micro = 0; double totalArea = 0.0;
                foreach (var fo in faces)
                {
                    var f = fo as IFace2; if (f == null) continue;
                    double a = 0.0; try { a = f.GetArea(); } catch { }
                    totalArea += a;
                    if (a > 0 && a <= MicroFaceAreaTol) micro++;
                }
                // If overall area is tiny, downgrade microface counts to avoid false positives on small parts
                if (totalArea > 0 && totalArea < 1e-5) // < 10 mm^2 total
                {
                    micro = 0;
                }
                res.MicroFaceCount = micro;
            }
            catch { }
        }

        private static double GetEdgeChordLength(IEdge edge)
        {
            try
            {
                var sv = edge.GetStartVertex() as IVertex;
                var ev = edge.GetEndVertex() as IVertex;
                var sa = (sv?.GetPoint() as object) as double[];
                var ea = (ev?.GetPoint() as object) as double[];
                if (sa != null && ea != null && sa.Length >= 3 && ea.Length >= 3)
                {
                    double dx = ea[0] - sa[0];
                    double dy = ea[1] - sa[1];
                    double dz = ea[2] - sa[2];
                    return Math.Sqrt(dx * dx + dy * dy + dz * dz);
                }
            }
            catch { }
            return 0.0;
        }

        private static bool InvokeCurveBool(ICurve c, string methodName, bool defaultValue = false)
        {
            try
            {
                if (c == null) return defaultValue;
                var mi = c.GetType().GetMethod(methodName);
                if (mi == null) return defaultValue;
                var r = mi.Invoke(c, null);
                if (r is bool b) return b;
            }
            catch { }
            return defaultValue;
        }

        private static string BuildReasonText(Result r)
        {
            if (!r.IsProblem) return string.Empty;
            // Compose a concise and human-friendly one-liner
            if (!string.IsNullOrEmpty(r.Reason)) return r.Reason;
            return "Failed preflight";
        }
    }
}
