using System.Collections.Generic;
using NM.Core.Pdf;

namespace NM.Core.Pdf.Models
{
    /// <summary>
    /// A GD&amp;T feature control frame extracted from a drawing.
    /// </summary>
    public sealed class GdtCallout
    {
        /// <summary>GD&amp;T type (Position, Flatness, Parallelism, etc.).</summary>
        public GdtType GdtFeatureType { get; set; }

        /// <summary>Legacy string type name for backward compatibility.</summary>
        public string Type { get; set; }

        /// <summary>Tolerance value as string (e.g., "0.005").</summary>
        public string Tolerance { get; set; }

        /// <summary>Tolerance value parsed as double (inches). Null if unparseable.</summary>
        public double? ToleranceValue { get; set; }

        /// <summary>True if tolerance applies at MMC (Maximum Material Condition).</summary>
        public bool IsMmc { get; set; }

        /// <summary>True if tolerance applies at LMC (Least Material Condition).</summary>
        public bool IsLmc { get; set; }

        /// <summary>True if preceded by diameter symbol (diametral tolerance zone).</summary>
        public bool IsDiametral { get; set; }

        /// <summary>Datum references (e.g., A, B, C).</summary>
        public List<string> DatumReferences { get; } = new List<string>();

        /// <summary>Description of the controlled feature (if identifiable).</summary>
        public string FeatureDescription { get; set; }

        /// <summary>Cost impact tier.</summary>
        public ToleranceTier Tier { get; set; }

        /// <summary>Cost impact level.</summary>
        public CostImpact Impact { get; set; }

        /// <summary>Raw text from the drawing.</summary>
        public string RawText { get; set; }

        /// <summary>Extraction confidence (0.0-1.0).</summary>
        public double Confidence { get; set; }
    }

    /// <summary>
    /// GD&amp;T geometric tolerance types per ASME Y14.5.
    /// </summary>
    public enum GdtType
    {
        // Form tolerances (no datum required)
        Flatness,
        Straightness,
        Circularity,
        Cylindricity,

        // Orientation tolerances (datum required)
        Parallelism,
        Perpendicularity,
        Angularity,

        // Location tolerances (datum required)
        Position,
        Concentricity,
        Symmetry,

        // Profile tolerances
        ProfileOfLine,
        ProfileOfSurface,

        // Runout tolerances (datum required)
        CircularRunout,
        TotalRunout
    }
}
