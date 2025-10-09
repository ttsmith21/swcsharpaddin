using System;
using NM.Core.Models;
using NM.Core.Validation;
using NM.StepClassifierAddin.Classification;
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
            if (res == null) { modelInfo.MarkValidated(false, "Analysis failed"); ValidationStats.Record(false); return ValidationResult.Fail("Analysis failed"); }

            if (res.IsProblem)
            {
                var reason = string.IsNullOrWhiteSpace(res.Reason) ? "Preflight failed" : res.Reason;
                modelInfo.MarkValidated(false, reason);
                ValidationStats.Record(false);
                return ValidationResult.Fail(reason);
            }

            modelInfo.MarkValidated(true);
            ValidationStats.Record(true);
            return ValidationResult.Ok();
        }
    }
}
