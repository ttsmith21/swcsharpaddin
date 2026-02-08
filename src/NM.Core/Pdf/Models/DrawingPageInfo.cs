using System.Collections.Generic;
using NM.Core.Pdf;

namespace NM.Core.Pdf.Models
{
    /// <summary>
    /// Information extracted from a single page/sheet of a PDF drawing.
    /// A multi-page PDF produces one DrawingPageInfo per page.
    /// </summary>
    public sealed class DrawingPageInfo
    {
        /// <summary>Path to the source PDF file.</summary>
        public string PdfPath { get; set; }

        /// <summary>1-based page number within the PDF.</summary>
        public int PageNumber { get; set; }

        /// <summary>Part number extracted from this page's title block.</summary>
        public string PartNumber { get; set; }

        /// <summary>Description from this page's title block.</summary>
        public string Description { get; set; }

        /// <summary>Revision from this page's title block.</summary>
        public string Revision { get; set; }

        /// <summary>Material from this page's title block.</summary>
        public string Material { get; set; }

        /// <summary>Sheet info (e.g., "1 OF 3", "SHEET 2").</summary>
        public string SheetInfo { get; set; }

        /// <summary>True if this page appears to be an assembly-level drawing.</summary>
        public bool IsAssemblyLevel { get; set; }

        /// <summary>True if a BOM table was detected on this page.</summary>
        public bool HasBom { get; set; }

        /// <summary>BOM entries found on this page (if any).</summary>
        public List<BomEntry> BomEntries { get; } = new List<BomEntry>();

        /// <summary>Manufacturing notes found on this page.</summary>
        public List<DrawingNote> Notes { get; } = new List<DrawingNote>();

        /// <summary>Routing hints derived from notes on this page.</summary>
        public List<RoutingHint> RoutingHints { get; } = new List<RoutingHint>();

        /// <summary>Recognized formal specifications (ASTM, AMS, MIL-SPEC, etc.).</summary>
        public List<SpecMatch> RecognizedSpecs { get; } = new List<SpecMatch>();

        /// <summary>GD&amp;T callouts extracted from this page.</summary>
        public List<GdtCallout> GdtCallouts { get; } = new List<GdtCallout>();

        /// <summary>Tolerance analysis results for this page.</summary>
        public ToleranceAnalysisResult ToleranceAnalysis { get; set; }

        /// <summary>Confidence in the title block extraction (0.0 - 1.0).</summary>
        public double Confidence { get; set; }

        /// <summary>True if this page had extractable text.</summary>
        public bool HasText { get; set; }
    }
}
