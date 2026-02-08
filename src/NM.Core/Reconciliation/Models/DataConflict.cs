namespace NM.Core.Reconciliation.Models
{
    /// <summary>
    /// A conflict where the 3D model and 2D drawing disagree on a field value.
    /// Requires human review to resolve.
    /// </summary>
    public sealed class DataConflict
    {
        /// <summary>The field name (e.g., "Material", "Thickness").</summary>
        public string Field { get; set; }

        /// <summary>Value from the 3D model.</summary>
        public string ModelValue { get; set; }

        /// <summary>Value from the PDF drawing.</summary>
        public string DrawingValue { get; set; }

        /// <summary>Which source the system recommends.</summary>
        public ConflictResolution Recommendation { get; set; }

        /// <summary>Why the system recommends this resolution.</summary>
        public string Reason { get; set; }

        /// <summary>Severity level.</summary>
        public ConflictSeverity Severity { get; set; }
    }

    public enum ConflictResolution
    {
        /// <summary>Use the 3D model value.</summary>
        UseModel,

        /// <summary>Use the drawing value.</summary>
        UseDrawing,

        /// <summary>Cannot determine — human must decide.</summary>
        HumanRequired
    }

    public enum ConflictSeverity
    {
        /// <summary>Minor (formatting/case difference). Auto-resolvable.</summary>
        Low,

        /// <summary>Moderate (different but plausibly equivalent values).</summary>
        Medium,

        /// <summary>Critical (completely different values). Must review.</summary>
        High
    }
}
