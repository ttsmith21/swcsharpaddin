namespace NM.Core.Reconciliation.Models
{
    /// <summary>
    /// A field that was missing in the 3D model but found in the PDF drawing.
    /// </summary>
    public sealed class GapFill
    {
        /// <summary>The field name (e.g., "Description", "PartNumber").</summary>
        public string Field { get; set; }

        /// <summary>The value extracted from the drawing.</summary>
        public string Value { get; set; }

        /// <summary>Where the value came from.</summary>
        public string Source { get; set; }

        /// <summary>Confidence in the extracted value (0.0 - 1.0).</summary>
        public double Confidence { get; set; }

        /// <summary>Whether to auto-apply (confidence > 0.85) or require human review.</summary>
        public bool AutoApply => Confidence >= 0.85;
    }
}
