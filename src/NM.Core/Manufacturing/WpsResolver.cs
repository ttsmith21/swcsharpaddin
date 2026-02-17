using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace NM.Core.Manufacturing
{
    /// <summary>
    /// Resolves the correct WPS for a weld joint based on materials, thickness, and joint type.
    /// Flags joints that need human review (thick groove welds, dissimilar metals, etc.).
    ///
    /// Business rules (from Northern Manufacturing):
    /// - Groove weld thickness > 0.5" → needs human review
    /// - Dissimilar base metals (e.g., CS to SS) → needs human review
    /// - Aluminum welds → always flag for verification
    /// - No matching WPS → flag as needing qualification
    /// </summary>
    public static class WpsResolver
    {
        private const double ThickGrooveThresholdIn = 0.5; // 1/2"

        /// <summary>
        /// Resolves WPS for a single weld joint.
        /// </summary>
        /// <param name="input">Joint description (materials, thickness, joint type).</param>
        /// <param name="table">Loaded WPS lookup table.</param>
        /// <returns>Match result with WPS selection and review flags.</returns>
        public static WpsMatchResult Resolve(WpsJointInput input, WpsLookupTable table)
        {
            var result = new WpsMatchResult();
            if (input == null)
            {
                result.Summary = "No joint input provided";
                return result;
            }

            string m1 = (input.BaseMetal1 ?? string.Empty).Trim().ToUpperInvariant();
            string m2 = (input.BaseMetal2 ?? string.Empty).Trim().ToUpperInvariant();

            // Flag: Dissimilar metals
            if (!string.IsNullOrEmpty(m1) && !string.IsNullOrEmpty(m2) && m1 != m2)
            {
                result.ReviewFlags.Add(new WpsReviewFlag(
                    WpsReviewReason.DissimilarMetals,
                    $"Dissimilar base metals: {m1} to {m2} — requires engineer review"));
            }

            // Flag: Thick groove weld (> 1/2")
            string jt = (input.JointType ?? string.Empty).Trim().ToUpperInvariant();
            if (jt == "GROOVE" && input.ThicknessIn > ThickGrooveThresholdIn)
            {
                result.ReviewFlags.Add(new WpsReviewFlag(
                    WpsReviewReason.ThickGrooveWeld,
                    $"Groove weld on {input.ThicknessIn:F3}\" material (>{ThickGrooveThresholdIn}\" threshold) — requires engineer review"));
            }

            // Flag: Aluminum welding
            if (m1 == "AL" || m2 == "AL")
            {
                result.ReviewFlags.Add(new WpsReviewFlag(
                    WpsReviewReason.AluminumWeld,
                    "Aluminum weld — verify procedure qualification and filler metal selection"));
            }

            // Look up matching WPS entries
            if (table != null && table.IsLoaded)
            {
                var matches = table.FindMatches(input);
                result.MatchedEntries.AddRange(matches);

                if (matches.Count == 0)
                {
                    result.ReviewFlags.Add(new WpsReviewFlag(
                        WpsReviewReason.NoMatchingWps,
                        $"No qualified WPS found for {m1}+{m2}, {input.ThicknessIn:F3}\", {input.JointType ?? "any"} joint"));
                }
                else if (matches.Count > 1)
                {
                    result.ReviewFlags.Add(new WpsReviewFlag(
                        WpsReviewReason.AmbiguousMatch,
                        $"{matches.Count} WPS entries match — engineer should confirm: {string.Join(", ", matches.Select(e => e.WpsNumber))}"));
                }
            }
            else
            {
                // No table loaded — can still flag review conditions but can't resolve WPS
                result.ReviewFlags.Add(new WpsReviewFlag(
                    WpsReviewReason.NoMatchingWps,
                    "WPS lookup table not loaded — cannot resolve procedure"));
            }

            // Build summary
            result.Summary = BuildSummary(result, input);

            return result;
        }

        /// <summary>
        /// Resolves WPS for a part based on its own material and thickness.
        /// For single-part processing, assumes a similar-metal joint (material welded to itself).
        /// The joint type can be provided from drawing weld symbols.
        /// </summary>
        /// <param name="materialCategory">Material category: "CS", "SS", "AL".</param>
        /// <param name="thicknessIn">Part thickness in inches.</param>
        /// <param name="jointType">Joint type from drawing: "Groove", "Fillet", or empty.</param>
        /// <param name="table">Loaded WPS lookup table.</param>
        /// <returns>Match result.</returns>
        public static WpsMatchResult ResolveForPart(string materialCategory, double thicknessIn,
            string jointType, WpsLookupTable table)
        {
            var input = new WpsJointInput
            {
                BaseMetal1 = materialCategory,
                BaseMetal2 = materialCategory, // assume similar joint
                ThicknessIn = thicknessIn,
                JointType = jointType
            };
            return Resolve(input, table);
        }

        /// <summary>
        /// Resolves WPS for an assembly joint between two components.
        /// Uses the thinner member's thickness as the governing thickness.
        /// </summary>
        public static WpsMatchResult ResolveForAssemblyJoint(
            string material1Category, double thickness1In,
            string material2Category, double thickness2In,
            string jointType, WpsLookupTable table)
        {
            // Governing thickness = thinner member for groove welds
            double govThickness = Math.Min(thickness1In, thickness2In);
            if (govThickness <= 0)
                govThickness = Math.Max(thickness1In, thickness2In);

            var input = new WpsJointInput
            {
                BaseMetal1 = material1Category,
                BaseMetal2 = material2Category,
                ThicknessIn = govThickness,
                JointType = jointType
            };
            return Resolve(input, table);
        }

        private static string BuildSummary(WpsMatchResult result, WpsJointInput input)
        {
            var parts = new List<string>();
            string metals = $"{input.BaseMetal1 ?? "?"} + {input.BaseMetal2 ?? "?"}";
            string thickness = input.ThicknessIn > 0
                ? $"{input.ThicknessIn:F3}\""
                : "unknown thickness";

            if (result.HasMatch)
            {
                parts.Add($"WPS {result.WpsNumber} ({result.MatchedEntries[0].Process})");
                parts.Add($"for {metals} {input.JointType ?? "weld"} at {thickness}");
            }
            else
            {
                parts.Add($"No WPS found for {metals} {input.JointType ?? "weld"} at {thickness}");
            }

            if (result.NeedsReview)
            {
                parts.Add($"— REVIEW: {string.Join("; ", result.ReviewFlags.Select(f => f.Reason))}");
            }

            return string.Join(" ", parts);
        }
    }
}
