using System;
using NM.Core.Models;

namespace NM.Core.Validation
{
    /// <summary>
    /// Orchestrates part validation using existing geometry checks.
    /// </summary>
    public sealed class PartValidator
    {
        // TODO: replace placeholder with a thin adapter to the actual body geometry validator
        public ValidationResult Validate(SwModelInfo modelInfo, object partDoc)
        {
            if (modelInfo == null) throw new ArgumentNullException(nameof(modelInfo));
            if (partDoc == null) return ValidationResult.Fail("No part document");
            if (modelInfo.Type != SwModelInfo.ModelType.Part) return ValidationResult.Fail($"Expected part, got {modelInfo.Type}");

            // For now, mark as valid to enable pipeline scaffolding.
            modelInfo.MarkValidated(true);
            return ValidationResult.Ok();
        }
    }
}
