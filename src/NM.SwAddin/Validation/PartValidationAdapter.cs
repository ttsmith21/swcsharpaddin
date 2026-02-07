using System;
using System.IO;
using NM.Core.Models;
using NM.Core.Processing;
using NM.Core.Validation;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Validation
{
    /// <summary>
    /// Adapter that runs SolidWorks-side preflight and returns a Core ValidationResult without exposing COM to NM.Core.
    /// Also records session pass/fail counts.
    /// </summary>
    public sealed class PartValidationAdapter
    {
        /// <summary>
        /// Validates a model and returns the result. Does NOT modify model state -
        /// the caller (BatchValidator) is responsible for state transitions.
        /// </summary>
        public ValidationResult Validate(SwModelInfo modelInfo, object modelDoc)
        {
            if (modelInfo == null) throw new ArgumentNullException(nameof(modelInfo));
            if (modelDoc == null) { ValidationStats.Record(false); return ValidationResult.Fail("No document"); }

            var model = modelDoc as IModelDoc2;
            if (model == null) { ValidationStats.Record(false); return ValidationResult.Fail("Invalid model handle"); }

            int t = (int)swDocumentTypes_e.swDocNONE;
            try { t = model.GetType(); } catch { }
            if (t != (int)swDocumentTypes_e.swDocPART)
            { ValidationStats.Record(false); return ValidationResult.Fail($"Expected part, got {(swDocumentTypes_e)t}"); }

            var res = PartPreflight.Analyze(model);
            if (res == null) { ValidationStats.Record(false); return ValidationResult.Fail("Analysis failed"); }

            // Run purchased-part heuristics on the geometry metrics
            string purchasedHint = null;
            if (res.FaceCount > 0 || res.MassKg > 0)
            {
                var hInput = new PurchasedPartHeuristics.HeuristicInput
                {
                    MassKg = res.MassKg,
                    FaceCount = res.FaceCount,
                    EdgeCount = res.EdgeCount,
                    FileName = Path.GetFileName(modelInfo.FilePath ?? string.Empty)
                };
                var hResult = PurchasedPartHeuristics.Analyze(hInput);
                if (hResult.LikelyPurchased)
                    purchasedHint = hResult.Reason;
            }

            if (res.IsProblem)
            {
                var reason = string.IsNullOrWhiteSpace(res.Reason) ? "Preflight failed" : res.Reason;
                ValidationStats.Record(false);
                var vr = ValidationResult.Fail(reason);
                vr.PurchasedHint = purchasedHint;
                vr.FaceCount = res.FaceCount;
                vr.EdgeCount = res.EdgeCount;
                vr.MassKg = res.MassKg;
                return vr;
            }

            ValidationStats.Record(true);
            var okResult = ValidationResult.Ok();
            okResult.PurchasedHint = purchasedHint;
            okResult.FaceCount = res.FaceCount;
            okResult.EdgeCount = res.EdgeCount;
            okResult.MassKg = res.MassKg;
            return okResult;
        }
    }
}
