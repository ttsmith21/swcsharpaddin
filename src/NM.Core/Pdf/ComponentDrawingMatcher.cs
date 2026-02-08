using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NM.Core.Pdf.Models;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Matches assembly components to drawing pages in a DrawingPackageIndex.
    /// Tries multiple strategies: exact part number, file name, BOM reference, fuzzy match.
    /// Pure logic — no SolidWorks types.
    /// </summary>
    public sealed class ComponentDrawingMatcher
    {
        /// <summary>
        /// Attempts to match a component to drawing pages in the index.
        /// Tries strategies in order of confidence: exact → file name → BOM → fuzzy.
        /// </summary>
        /// <param name="componentPath">Full path to the component file.</param>
        /// <param name="partNumber">Part number from the component's custom properties (may be null).</param>
        /// <param name="index">The drawing package index to search.</param>
        /// <returns>Match result with the drawing pages and match method.</returns>
        public ComponentMatch Match(string componentPath, string partNumber, DrawingPackageIndex index)
        {
            if (index == null)
                return ComponentMatch.NoMatch;

            // Strategy 1: Exact part number match
            if (!string.IsNullOrWhiteSpace(partNumber))
            {
                var pages = index.FindPages(partNumber);
                if (pages.Count > 0)
                {
                    return new ComponentMatch
                    {
                        Pages = pages,
                        Method = MatchMethod.ExactPartNumber,
                        Confidence = 0.95
                    };
                }
            }

            // Strategy 2: File name match (common when file is named after part number)
            if (!string.IsNullOrEmpty(componentPath))
            {
                string fileName = Path.GetFileNameWithoutExtension(componentPath);
                if (!string.IsNullOrEmpty(fileName))
                {
                    var pages = index.FindPages(fileName);
                    if (pages.Count > 0)
                    {
                        return new ComponentMatch
                        {
                            Pages = pages,
                            Method = MatchMethod.FileName,
                            Confidence = 0.85
                        };
                    }
                }
            }

            // Strategy 3: BOM reference (part number appears in a BOM table)
            if (!string.IsNullOrWhiteSpace(partNumber))
            {
                var bomMatch = FindViaBom(partNumber, index);
                if (bomMatch != null)
                    return bomMatch;
            }

            // Strategy 4: File name in BOM
            if (!string.IsNullOrEmpty(componentPath))
            {
                string fileName = Path.GetFileNameWithoutExtension(componentPath);
                if (!string.IsNullOrEmpty(fileName))
                {
                    var bomMatch = FindViaBom(fileName, index);
                    if (bomMatch != null)
                        return bomMatch;
                }
            }

            return ComponentMatch.NoMatch;
        }

        /// <summary>
        /// Matches all components in a list against the index.
        /// Returns matched and unmatched component paths.
        /// </summary>
        public MatchResults MatchAll(
            IList<ComponentInfo> components, DrawingPackageIndex index)
        {
            var results = new MatchResults();

            if (components == null || index == null)
                return results;

            foreach (var comp in components)
            {
                var match = Match(comp.FilePath, comp.PartNumber, index);
                if (match.IsMatched)
                {
                    results.Matched[comp.FilePath] = match;
                }
                else
                {
                    results.Unmatched.Add(comp.FilePath);
                }
            }

            // Find drawing pages that weren't matched to any component
            var matchedPartNumbers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in results.Matched)
            {
                foreach (var page in kv.Value.Pages)
                {
                    if (!string.IsNullOrEmpty(page.PartNumber))
                        matchedPartNumbers.Add(page.PartNumber);
                }
            }

            foreach (var kv in index.PagesByPartNumber)
            {
                if (!matchedPartNumbers.Contains(kv.Key))
                {
                    results.UnmatchedDrawings.AddRange(kv.Value);
                }
            }

            // Unmatched pages (no part number) are always unmatched drawings
            results.UnmatchedDrawings.AddRange(index.UnmatchedPages);

            return results;
        }

        private ComponentMatch FindViaBom(string searchTerm, DrawingPackageIndex index)
        {
            string normalized = searchTerm.Trim().ToUpperInvariant();

            foreach (var bomEntry in index.AllBomEntries)
            {
                if (string.IsNullOrEmpty(bomEntry.PartNumber))
                    continue;

                string bomPn = bomEntry.PartNumber.Trim().ToUpperInvariant();
                if (string.Equals(bomPn, normalized, StringComparison.OrdinalIgnoreCase) ||
                    bomPn.Contains(normalized) || normalized.Contains(bomPn))
                {
                    // BOM confirms this part belongs in the assembly, but the drawing
                    // details may be on the assembly page, not a separate part page.
                    // Check if there's a dedicated part page first.
                    var pages = index.FindPages(bomEntry.PartNumber);
                    if (pages.Count > 0)
                    {
                        return new ComponentMatch
                        {
                            Pages = pages,
                            Method = MatchMethod.BomReference,
                            Confidence = 0.75
                        };
                    }
                }
            }

            return null;
        }
    }

    /// <summary>
    /// Info about a component for matching purposes (no SW types).
    /// </summary>
    public sealed class ComponentInfo
    {
        public string FilePath { get; set; }
        public string PartNumber { get; set; }
        public bool IsAssembly { get; set; }
        public int Quantity { get; set; }
    }

    /// <summary>
    /// Result of matching a single component to drawing pages.
    /// </summary>
    public sealed class ComponentMatch
    {
        public static readonly ComponentMatch NoMatch = new ComponentMatch
        {
            Pages = new List<DrawingPageInfo>(),
            Method = MatchMethod.None,
            Confidence = 0
        };

        public List<DrawingPageInfo> Pages { get; set; } = new List<DrawingPageInfo>();
        public MatchMethod Method { get; set; }
        public double Confidence { get; set; }
        public bool IsMatched => Pages.Count > 0 && Method != MatchMethod.None;
    }

    /// <summary>
    /// Results from matching all components against a drawing index.
    /// </summary>
    public sealed class MatchResults
    {
        /// <summary>Components successfully matched (keyed by file path).</summary>
        public Dictionary<string, ComponentMatch> Matched { get; }
            = new Dictionary<string, ComponentMatch>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Component paths that had no matching drawing.</summary>
        public List<string> Unmatched { get; } = new List<string>();

        /// <summary>Drawing pages that did not match any component.</summary>
        public List<DrawingPageInfo> UnmatchedDrawings { get; } = new List<DrawingPageInfo>();
    }
}
