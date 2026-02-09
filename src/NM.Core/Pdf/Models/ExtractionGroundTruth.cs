using System.Collections.Generic;

namespace NM.Core.Pdf.Models
{
    /// <summary>
    /// Ground truth data for a single engineering drawing, used to measure extraction accuracy.
    /// Store as JSON alongside test PDFs in tests/GoldStandard_Inputs/Pdf/.
    /// </summary>
    public sealed class ExtractionGroundTruth
    {
        /// <summary>PDF filename (relative to test directory).</summary>
        public string PdfFileName { get; set; }

        /// <summary>Drawing type: CAD-generated, scanned, GD&T-heavy, multi-page, non-standard.</summary>
        public string DrawingType { get; set; }

        // Title block ground truth
        public string PartNumber { get; set; }
        public string Description { get; set; }
        public string Revision { get; set; }
        public string Material { get; set; }
        public string Finish { get; set; }
        public string DrawnBy { get; set; }
        public string Scale { get; set; }
        public string SheetInfo { get; set; }
        public string ToleranceGeneral { get; set; }

        // Manufacturing notes (exact text of each note on the drawing)
        public List<string> ManufacturingNotes { get; set; } = new List<string>();

        // GD&T callouts
        public List<GroundTruthGdt> GdtCallouts { get; set; } = new List<GroundTruthGdt>();

        // Tight tolerances (tolerances tighter than general tolerance)
        public List<GroundTruthTolerance> TightTolerances { get; set; } = new List<GroundTruthTolerance>();

        // Geometry
        public double? Thickness_in { get; set; }
        public double? OverallLength_in { get; set; }
        public double? OverallWidth_in { get; set; }

        /// <summary>Number of pages in the PDF.</summary>
        public int PageCount { get; set; } = 1;
    }

    public sealed class GroundTruthGdt
    {
        public string Type { get; set; }
        public string Tolerance { get; set; }
        public List<string> DatumReferences { get; set; } = new List<string>();
        public string FeatureDescription { get; set; }
    }

    public sealed class GroundTruthTolerance
    {
        public string DimensionValue { get; set; }
        public string Tolerance { get; set; }
        public string FeatureDescription { get; set; }
    }
}
