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
                var bodies = part.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
                if (bodies == null || bodies.Length == 0)
                    return new PreflightResult { IsProblem = true, Reason = "No solid bodies found" };

                // Multi-body check - FAIL validation for multi-body parts
                if (bodies.Length > 1)
                {
                    return new PreflightResult { IsProblem = true, Reason = $"Multi-body part ({bodies.Length} bodies)" };
                }

                // Material check - FAIL validation if no material assigned
                var material = SolidWorksApiWrapper.GetMaterialName(model);
                if (string.IsNullOrWhiteSpace(material))
                {
                    return new PreflightResult { IsProblem = true, Reason = "No material assigned" };
                }

                return new PreflightResult { IsProblem = false, Reason = null };
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
    }
}
