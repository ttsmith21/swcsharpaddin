using System.Collections.Generic;
using System.Linq;

namespace NM.Core.Pdf.Models
{
    /// <summary>
    /// Result of batch analysis across an entire drawing package / assembly.
    /// </summary>
    public sealed class BatchAnalysisResult
    {
        /// <summary>Name of the top-level assembly or file being analyzed.</summary>
        public string TopLevelName { get; set; }

        /// <summary>The drawing package index used for matching.</summary>
        public DrawingPackageIndex PackageIndex { get; set; }

        /// <summary>Per-component results (keyed by file path, case-insensitive).</summary>
        public Dictionary<string, ComponentAnalysisResult> ComponentResults { get; }
            = new Dictionary<string, ComponentAnalysisResult>(System.StringComparer.OrdinalIgnoreCase);

        /// <summary>Component paths that had no matching drawing page.</summary>
        public List<string> UnmatchedComponents { get; } = new List<string>();

        /// <summary>Drawing pages that did not match any component.</summary>
        public List<DrawingPageInfo> UnmatchedDrawings { get; } = new List<DrawingPageInfo>();

        /// <summary>Components that were skipped (suppressed, virtual, etc.).</summary>
        public List<string> SkippedComponents { get; } = new List<string>();

        /// <summary>Total unique components in the assembly.</summary>
        public int TotalComponents { get; set; }

        /// <summary>Components successfully matched to drawing pages.</summary>
        public int MatchedComponents => ComponentResults.Count;

        /// <summary>Components with at least one property suggestion.</summary>
        public int ComponentsWithSuggestions =>
            ComponentResults.Values.Count(r => r.SuggestionCount > 0);

        /// <summary>Total property suggestions across all components.</summary>
        public int TotalSuggestions =>
            ComponentResults.Values.Sum(r => r.SuggestionCount);

        /// <summary>Human-readable summary.</summary>
        public string Summary =>
            $"{MatchedComponents}/{TotalComponents} components matched, " +
            $"{TotalSuggestions} suggestions, " +
            $"{UnmatchedComponents.Count} unmatched parts, " +
            $"{UnmatchedDrawings.Count} unmatched drawings";
    }

    /// <summary>
    /// Analysis result for a single component in the assembly.
    /// </summary>
    public sealed class ComponentAnalysisResult
    {
        /// <summary>Full path to the component file.</summary>
        public string FilePath { get; set; }

        /// <summary>File name without extension.</summary>
        public string FileName { get; set; }

        /// <summary>Part number extracted from the component's custom properties.</summary>
        public string PartNumber { get; set; }

        /// <summary>Whether this component is an assembly (vs a part).</summary>
        public bool IsAssembly { get; set; }

        /// <summary>Quantity of this component in the parent assembly.</summary>
        public int Quantity { get; set; }

        /// <summary>Drawing data extracted from the matched PDF page(s).</summary>
        public DrawingData DrawingData { get; set; }

        /// <summary>Number of property suggestions generated.</summary>
        public int SuggestionCount { get; set; }

        /// <summary>Number of conflicts found during reconciliation.</summary>
        public int ConflictCount { get; set; }

        /// <summary>Whether a file rename was suggested.</summary>
        public bool HasRenameSuggestion { get; set; }

        /// <summary>Overall confidence in the match (0.0-1.0).</summary>
        public double MatchConfidence { get; set; }

        /// <summary>How the component was matched to its drawing page(s).</summary>
        public MatchMethod MatchMethod { get; set; }
    }

    /// <summary>
    /// How a component was matched to drawing pages.
    /// </summary>
    public enum MatchMethod
    {
        /// <summary>No match found.</summary>
        None,

        /// <summary>Exact part number match.</summary>
        ExactPartNumber,

        /// <summary>File name matched a part number in the drawing index.</summary>
        FileName,

        /// <summary>Partial/fuzzy part number match.</summary>
        FuzzyPartNumber,

        /// <summary>Matched via BOM entry on an assembly drawing.</summary>
        BomReference
    }
}
