using System;
using System.Collections.Generic;
using NM.Core.Pdf;

namespace NM.Core.Pdf.Models
{
    /// <summary>
    /// Central DTO for all data extracted from a PDF engineering drawing.
    /// The PDF equivalent of PartData — populated by PdfDrawingAnalyzer.
    /// </summary>
    public sealed class DrawingData
    {
        // Title block fields
        public string PartNumber { get; set; }
        public string Description { get; set; }
        public string Revision { get; set; }
        public string Material { get; set; }
        public string Finish { get; set; }
        public string DrawnBy { get; set; }
        public string CheckedBy { get; set; }
        public string ApprovedBy { get; set; }
        public DateTime? DrawingDate { get; set; }
        public string Scale { get; set; }
        public string SheetInfo { get; set; }  // e.g., "1 OF 3"
        public string ToleranceGeneral { get; set; }

        // Geometry hints extracted from dimensions/notes
        public double? Thickness_in { get; set; }
        public double? OverallLength_in { get; set; }
        public double? OverallWidth_in { get; set; }

        // Manufacturing notes
        public List<DrawingNote> Notes { get; } = new List<DrawingNote>();

        // GD&T callouts
        public List<GdtCallout> GdtCallouts { get; } = new List<GdtCallout>();

        // BOM entries (assembly drawings)
        public List<BomEntry> BomEntries { get; } = new List<BomEntry>();

        // Routing hints derived from notes
        public List<RoutingHint> RoutingHints { get; } = new List<RoutingHint>();

        // Recognized specifications (ASTM, AMS, MIL-SPEC, AWS, etc.)
        public List<SpecMatch> RecognizedSpecs { get; } = new List<SpecMatch>();

        // Source tracking
        public string SourcePdfPath { get; set; }
        public int PageCount { get; set; }
        public AnalysisMethod Method { get; set; }
        public double OverallConfidence { get; set; }

        // Raw extracted text (for debugging / AI fallback)
        public string RawText { get; set; }
    }

    public enum AnalysisMethod
    {
        TextOnly,      // PdfPig text extraction + regex
        VisionAI,      // Claude/GPT-4o vision API
        Hybrid         // Text extraction + AI for gaps
    }
}
