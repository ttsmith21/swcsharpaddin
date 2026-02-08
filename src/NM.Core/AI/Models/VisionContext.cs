using System.Collections.Generic;

namespace NM.Core.AI.Models
{
    /// <summary>
    /// Context information provided to the AI to improve analysis accuracy.
    /// Pass known data from the 3D model so the AI can cross-reference.
    /// </summary>
    public sealed class VisionContext
    {
        /// <summary>Known part number from the 3D model (for cross-validation).</summary>
        public string KnownPartNumber { get; set; }

        /// <summary>Known material from the 3D model.</summary>
        public string KnownMaterial { get; set; }

        /// <summary>Known thickness from the 3D model (inches).</summary>
        public double? KnownThickness_in { get; set; }

        /// <summary>Known part classification (SheetMetal, Tube, etc.).</summary>
        public string KnownClassification { get; set; }

        /// <summary>Company name (helps identify title block format).</summary>
        public string CompanyName { get; set; }

        /// <summary>Whether this is a title-block-only analysis.</summary>
        public bool TitleBlockOnly { get; set; }

        /// <summary>Additional context hints for the AI.</summary>
        public List<string> Hints { get; } = new List<string>();

        /// <summary>True if any context data is available.</summary>
        public bool HasContext =>
            !string.IsNullOrEmpty(KnownPartNumber) ||
            !string.IsNullOrEmpty(KnownMaterial) ||
            KnownThickness_in.HasValue;
    }
}
