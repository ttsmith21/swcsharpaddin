using System;

namespace NM.Core.Validation
{
    /// <summary>
    /// Simple validation result for pipeline steps.
    /// </summary>
    public sealed class ValidationResult
    {
        public bool Success { get; }
        public string Summary { get; }

        private ValidationResult(bool success, string summary)
        {
            Success = success;
            Summary = summary ?? string.Empty;
        }

        public static ValidationResult Ok(string summary = "Valid") => new ValidationResult(true, summary);
        public static ValidationResult Fail(string reason) => new ValidationResult(false, string.IsNullOrWhiteSpace(reason) ? "Failed" : reason);
    }
}
