using System;
using System.Collections.Generic;
using System.Linq;

namespace NM.Core.Pdf.Models
{
    /// <summary>
    /// Index of all drawing pages found in a drawing package (folder of PDFs or multi-page PDF).
    /// Maps part numbers to their drawing pages, tracks BOM entries, and identifies unmatched pages.
    /// </summary>
    public sealed class DrawingPackageIndex
    {
        /// <summary>All pages scanned, grouped by part number (case-insensitive).</summary>
        public Dictionary<string, List<DrawingPageInfo>> PagesByPartNumber { get; }
            = new Dictionary<string, List<DrawingPageInfo>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Pages where no part number could be extracted.</summary>
        public List<DrawingPageInfo> UnmatchedPages { get; } = new List<DrawingPageInfo>();

        /// <summary>Aggregated BOM entries from all assembly-level pages.</summary>
        public List<BomEntry> AllBomEntries { get; } = new List<BomEntry>();

        /// <summary>All PDF files that were scanned.</summary>
        public List<string> ScannedFiles { get; } = new List<string>();

        /// <summary>Total number of pages scanned across all PDFs.</summary>
        public int TotalPages { get; set; }

        /// <summary>Number of pages successfully matched to a part number.</summary>
        public int MatchedPages => PagesByPartNumber.Values.Sum(list => list.Count);

        /// <summary>Number of unique part numbers found.</summary>
        public int UniquePartNumbers => PagesByPartNumber.Count;

        /// <summary>
        /// Finds drawing pages for a given part number.
        /// Tries exact match first, then partial/fuzzy matching.
        /// </summary>
        public List<DrawingPageInfo> FindPages(string partNumber)
        {
            if (string.IsNullOrWhiteSpace(partNumber))
                return new List<DrawingPageInfo>();

            // Exact match (case-insensitive via dictionary comparer)
            if (PagesByPartNumber.TryGetValue(partNumber, out var pages))
                return pages;

            // Try trimmed/normalized match
            string normalized = partNumber.Trim().ToUpperInvariant();
            foreach (var kv in PagesByPartNumber)
            {
                if (string.Equals(kv.Key.Trim(), normalized, StringComparison.OrdinalIgnoreCase))
                    return kv.Value;
            }

            // Try partial match (part number contained in or containing the key)
            foreach (var kv in PagesByPartNumber)
            {
                string key = kv.Key.Trim().ToUpperInvariant();
                if (key.Contains(normalized) || normalized.Contains(key))
                    return kv.Value;
            }

            return new List<DrawingPageInfo>();
        }

        /// <summary>
        /// Builds a combined DrawingData from all pages matching a part number.
        /// Merges notes, routing hints, and BOM entries across sheets.
        /// </summary>
        public DrawingData BuildDrawingData(string partNumber)
        {
            var pages = FindPages(partNumber);
            if (pages.Count == 0)
                return null;

            var primary = pages[0]; // Use first page for title block data
            var result = new DrawingData
            {
                PartNumber = primary.PartNumber,
                Description = primary.Description,
                Revision = primary.Revision,
                Material = primary.Material,
                SheetInfo = primary.SheetInfo,
                SourcePdfPath = primary.PdfPath,
                PageCount = pages.Count,
                Method = AnalysisMethod.TextOnly,
                OverallConfidence = primary.Confidence
            };

            // Merge notes and routing hints from all pages
            var seenNotes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var page in pages)
            {
                foreach (var note in page.Notes)
                {
                    if (seenNotes.Add(note.Text))
                        result.Notes.Add(note);
                }
                foreach (var hint in page.RoutingHints)
                {
                    result.RoutingHints.Add(hint);
                }
                foreach (var bom in page.BomEntries)
                {
                    result.BomEntries.Add(bom);
                }
            }

            return result;
        }

        /// <summary>Human-readable summary of the scan results.</summary>
        public string Summary =>
            $"{ScannedFiles.Count} PDF(s), {TotalPages} pages, " +
            $"{UniquePartNumbers} part numbers found, " +
            $"{UnmatchedPages.Count} unmatched pages";
    }
}
