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

        /// <summary>
        /// Optional heuristic hint (e.g., "Likely purchased: Low mass; High face count").
        /// Set by the validation adapter when geometry metrics suggest a purchased part.
        /// </summary>
        public string PurchasedHint { get; set; }

        /// <summary>Geometry metrics from preflight for downstream heuristic analysis.</summary>
        public int FaceCount { get; set; }
        public int EdgeCount { get; set; }
        public double MassKg { get; set; }

        private ValidationResult(bool success, string summary)
        {
            Success = success;
            Summary = summary ?? string.Empty;
        }

        public static ValidationResult Ok(string summary = "Valid") => new ValidationResult(true, summary);
        public static ValidationResult Fail(string reason) => new ValidationResult(false, string.IsNullOrWhiteSpace(reason) ? "Failed" : reason);
    }
}
