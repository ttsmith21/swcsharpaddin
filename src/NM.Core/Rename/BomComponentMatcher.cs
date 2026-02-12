using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace NM.Core.Rename
{
    /// <summary>
    /// Matches BOM rows from a PDF drawing to assembly component filenames.
    /// Uses layered heuristics: exact match, substring, quantity correlation.
    /// Falls back to AI matching when heuristics are insufficient.
    /// </summary>
    public sealed class BomComponentMatcher
    {
        /// <summary>
        /// Match BOM rows to components using heuristic rules.
        /// Returns a RenameEntry per component with the best BOM match (if any).
        /// </summary>
        public List<RenameEntry> Match(List<BomRow> bomRows, List<ComponentInfo> components)
        {
            if (bomRows == null) throw new ArgumentNullException(nameof(bomRows));
            if (components == null) throw new ArgumentNullException(nameof(components));

            var entries = new List<RenameEntry>();

            foreach (var comp in components)
            {
                string nameNoExt = Path.GetFileNameWithoutExtension(comp.FileName) ?? comp.FileName;

                var bestMatch = FindBestMatch(nameNoExt, comp.InstanceCount, bomRows);

                var entry = new RenameEntry
                {
                    Index = comp.Index,
                    CurrentFileName = comp.FileName,
                    CurrentFilePath = comp.FilePath,
                    Configuration = comp.Configuration,
                    MatchedBomRow = bestMatch.Row,
                    PredictedName = bestMatch.Row != null
                        ? SanitizeName(bestMatch.Row.PartNumber, bestMatch.Row.Description)
                        : nameNoExt,
                    FinalName = bestMatch.Row != null
                        ? SanitizeName(bestMatch.Row.PartNumber, bestMatch.Row.Description)
                        : nameNoExt,
                    Confidence = bestMatch.Confidence,
                    MatchReason = bestMatch.Reason,
                    IsApproved = bestMatch.Confidence >= 0.80
                };

                entries.Add(entry);
            }

            return entries;
        }

        /// <summary>
        /// Apply AI matching results to existing entries.
        /// Called after the AI fallback returns its JSON response.
        /// </summary>
        public void ApplyAiMatches(
            List<RenameEntry> entries,
            List<BomRow> bomRows,
            List<AiMatchResult> aiMatches)
        {
            if (aiMatches == null) return;

            foreach (var aim in aiMatches)
            {
                if (aim.ComponentIndex < 0 || aim.ComponentIndex >= entries.Count)
                    continue;

                var entry = entries[aim.ComponentIndex];

                // Only upgrade if AI is more confident than heuristic
                if (aim.Confidence <= entry.Confidence)
                    continue;

                var bomRow = bomRows.FirstOrDefault(b => b.ItemNumber == aim.BomItemNumber);
                if (bomRow == null) continue;

                entry.MatchedBomRow = bomRow;
                entry.PredictedName = !string.IsNullOrWhiteSpace(aim.PredictedName)
                    ? aim.PredictedName
                    : SanitizeName(bomRow.PartNumber, bomRow.Description);
                entry.FinalName = entry.PredictedName;
                entry.Confidence = aim.Confidence;
                entry.MatchReason = aim.Reason ?? "AI match";
                entry.IsApproved = aim.Confidence >= 0.80;
            }
        }

        private MatchCandidate FindBestMatch(string componentName, int instanceCount, List<BomRow> bomRows)
        {
            var best = new MatchCandidate { Confidence = 0, Reason = "No match" };

            foreach (var bom in bomRows)
            {
                double score = 0;
                string reason = "";

                // Layer 1: Exact part number in filename
                if (!string.IsNullOrWhiteSpace(bom.PartNumber) &&
                    componentName.IndexOf(bom.PartNumber, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    score = 0.95;
                    reason = $"Part number \"{bom.PartNumber}\" found in filename";
                }
                // Layer 2: Exact description in filename
                else if (!string.IsNullOrWhiteSpace(bom.Description) &&
                         componentName.IndexOf(bom.Description, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    score = 0.85;
                    reason = $"Description \"{bom.Description}\" found in filename";
                }
                // Layer 3: Substring overlap — component name contains significant BOM words
                else
                {
                    double overlap = CalculateWordOverlap(componentName, bom.PartNumber, bom.Description);
                    if (overlap > 0.3)
                    {
                        score = 0.40 + (overlap * 0.40); // 0.40 - 0.80 range
                        reason = $"Word overlap {overlap:P0} with BOM item {bom.ItemNumber}";
                    }
                }

                // Layer 4: Quantity correlation boost
                if (score > 0 && instanceCount > 0 && bom.Quantity == instanceCount)
                {
                    score = Math.Min(score + 0.05, 1.0);
                    reason += $" + qty match ({instanceCount})";
                }

                if (score > best.Confidence)
                {
                    best = new MatchCandidate
                    {
                        Row = bom,
                        Confidence = score,
                        Reason = reason
                    };
                }
            }

            return best;
        }

        private static double CalculateWordOverlap(string componentName, string partNumber, string description)
        {
            var compWords = TokenizeForMatching(componentName);
            if (compWords.Count == 0) return 0;

            var bomWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var w in TokenizeForMatching(partNumber)) bomWords.Add(w);
            foreach (var w in TokenizeForMatching(description)) bomWords.Add(w);

            if (bomWords.Count == 0) return 0;

            int matches = compWords.Count(w => bomWords.Contains(w));
            return (double)matches / Math.Max(compWords.Count, bomWords.Count);
        }

        private static List<string> TokenizeForMatching(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return new List<string>();

            // Split on common delimiters: space, dash, underscore, dot, caret
            var tokens = input.Split(new[] { ' ', '-', '_', '.', '^', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);

            // Filter out noise tokens
            return tokens
                .Where(t => t.Length >= 2) // skip single chars
                .Where(t => !IsNoiseToken(t))
                .ToList();
        }

        private static bool IsNoiseToken(string token)
        {
            // Common STEP import noise words
            switch (token.ToUpperInvariant())
            {
                case "SLDPRT":
                case "SLDASM":
                case "STEP":
                case "STP":
                case "BODY":
                case "MOVE":
                case "COPY":
                case "IMPORT":
                case "IMPORTED":
                    return true;
                default:
                    return false;
            }
        }

        private static string SanitizeName(string partNumber, string description)
        {
            // Prefer part number, fall back to description
            string name = !string.IsNullOrWhiteSpace(partNumber)
                ? partNumber.Trim()
                : !string.IsNullOrWhiteSpace(description)
                    ? description.Trim()
                    : "Unknown";

            // Remove invalid filename chars
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c.ToString(), "");

            return name.Trim();
        }

        private struct MatchCandidate
        {
            public BomRow Row;
            public double Confidence;
            public string Reason;
        }
    }

    /// <summary>
    /// Result from the AI matching fallback.
    /// Parsed from the JSON response of GetBomMatchingPrompt.
    /// </summary>
    public sealed class AiMatchResult
    {
        public int ComponentIndex { get; set; }
        public int BomItemNumber { get; set; }
        public string PredictedName { get; set; }
        public double Confidence { get; set; }
        public string Reason { get; set; }
    }
}
