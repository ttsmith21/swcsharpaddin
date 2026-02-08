using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Validation
{
    /// <summary>
    /// Part preflight validation - checks part for basic requirements before processing.
    /// </summary>
    public static class PartPreflight
    {
        /// <summary>
        /// Analyzes a part model for processing readiness.
        /// </summary>
        public static PreflightResult Analyze(IModelDoc2 model)
        {
            try
            {
                if (model == null)
                    return new PreflightResult { IsProblem = true, Reason = "Model is null" };

                int docType = model.GetType();
                if (docType != (int)swDocumentTypes_e.swDocPART)
                    return new PreflightResult { IsProblem = true, Reason = $"Not a part document (type: {docType})" };

                var part = model as IPartDoc;
                if (part == null)
                    return new PreflightResult { IsProblem = true, Reason = "Could not cast to IPartDoc" };

                // Check for solid bodies
                var solidBodies = part.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
                if (solidBodies == null || solidBodies.Length == 0)
                {
                    // Check if there are surface bodies (sheet bodies) - provide clearer message
                    var surfaceBodies = part.GetBodies2((int)swBodyType_e.swSheetBody, true) as object[];
                    if (surfaceBodies != null && surfaceBodies.Length > 0)
                    {
                        return new PreflightResult { IsProblem = true, Reason = $"Surface body only ({surfaceBodies.Length} surface bodies, no solid)" };
                    }
                    return new PreflightResult { IsProblem = true, Reason = "Empty part (no solid or surface bodies)" };
                }
                var bodies = solidBodies;

                // Multi-body check - FAIL validation for multi-body parts
                if (bodies.Length > 1)
                {
                    return new PreflightResult { IsProblem = true, Reason = $"Multi-body part ({bodies.Length} bodies)", SolidBodyCount = bodies.Length };
                }

                // Mixed-body check: solid + surface bodies coexisting
                // Parts like D8_Swagelock_T_Fitting have both solid and surface bodies.
                // Surface bodies can interfere with geometry analysis.
                var mixedSurfBodies = part.GetBodies2((int)swBodyType_e.swSheetBody, true) as object[];
                int surfaceCount = mixedSurfBodies?.Length ?? 0;
                if (surfaceCount > 0)
                {
                    return new PreflightResult
                    {
                        IsProblem = true,
                        Reason = $"Mixed-body part ({bodies.Length} solid + {surfaceCount} surface bodies)",
                        SolidBodyCount = bodies.Length,
                        SurfaceBodyCount = surfaceCount
                    };
                }

                // Material check - FAIL validation if no material assigned
                var material = SolidWorksApiWrapper.GetMaterialName(model);
                if (string.IsNullOrWhiteSpace(material))
                {
                    return new PreflightResult { IsProblem = true, Reason = "No material assigned", SolidBodyCount = bodies.Length };
                }

                // Extract geometry metrics for heuristic analysis
                var mainBody = (IBody2)bodies[0];
                var faces = mainBody.GetFaces() as object[];
                var edges = mainBody.GetEdges() as object[];

                var result = new PreflightResult
                {
                    IsProblem = false,
                    SolidBodyCount = bodies.Length,
                    FaceCount = faces?.Length ?? 0,
                    EdgeCount = edges?.Length ?? 0
                };

                try
                {
                    var massProp = model.Extension?.CreateMassProperty() as IMassProperty;
                    if (massProp != null)
                        result.MassKg = massProp.Mass;
                }
                catch { /* Non-fatal - mass is optional for heuristics */ }

                // Extract bounding box for oversize pre-screening and heuristics
                try
                {
                    double[] box = (double[])mainBody.GetBodyBox();
                    if (box != null && box.Length >= 6)
                    {
                        double dx = Math.Abs(box[3] - box[0]);
                        double dy = Math.Abs(box[4] - box[1]);
                        double dz = Math.Abs(box[5] - box[2]);
                        result.BBoxMaxDimM = Math.Max(dx, Math.Max(dy, dz));
                        result.BBoxMinDimM = Math.Min(dx, Math.Min(dy, dz));
                    }
                }
                catch { /* Non-fatal - bbox is optional for heuristics */ }

                return result;
            }
            catch (Exception ex)
            {
                return new PreflightResult { IsProblem = true, Reason = "Preflight exception: " + ex.Message };
            }
        }
    }

    /// <summary>
    /// Result of preflight analysis.
    /// </summary>
    public class PreflightResult
    {
        public bool IsProblem { get; set; }
        public string Reason { get; set; }

        // Geometry metrics for heuristic analysis (populated on success or partial success)
        public int SolidBodyCount { get; set; }
        public int SurfaceBodyCount { get; set; }
        public int FaceCount { get; set; }
        public int EdgeCount { get; set; }
        public double MassKg { get; set; }
        public double BBoxMaxDimM { get; set; }
        public double BBoxMinDimM { get; set; }
    }
}
