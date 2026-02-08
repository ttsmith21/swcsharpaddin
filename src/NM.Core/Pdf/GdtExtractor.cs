using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using NM.Core.Pdf.Models;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Extracts GD&amp;T (Geometric Dimensioning &amp; Tolerancing) feature control frames
    /// from PDF drawing text. Handles both Unicode symbols and text abbreviations since
    /// PDF text extraction often loses special GD&amp;T symbols.
    ///
    /// Complements ToleranceAnalyzer — that handles normal dimension tolerances (±0.005),
    /// this handles geometric tolerances (true position, flatness, parallelism, etc.).
    /// </summary>
    public sealed class GdtExtractor
    {
        // Map of GD&T type keywords/symbols to enum values
        private static readonly (string Pattern, GdtType Type, string DisplayName)[] GdtPatterns =
        {
            // Position (most common, most cost-impacting)
            (@"TRUE\s*POS(?:ITION)?", GdtType.Position, "Position"),
            (@"T/?P", GdtType.Position, "Position"),
            (@"⌖", GdtType.Position, "Position"),

            // Form tolerances
            (@"FLATNESS", GdtType.Flatness, "Flatness"),
            (@"⏥", GdtType.Flatness, "Flatness"),
            (@"STRAIGHTNESS", GdtType.Straightness, "Straightness"),
            (@"⏤", GdtType.Straightness, "Straightness"),
            (@"CIRCULARITY", GdtType.Circularity, "Circularity"),
            (@"ROUNDNESS", GdtType.Circularity, "Circularity"),
            (@"○", GdtType.Circularity, "Circularity"),
            (@"CYLINDRICITY", GdtType.Cylindricity, "Cylindricity"),
            (@"⌭", GdtType.Cylindricity, "Cylindricity"),

            // Orientation tolerances
            (@"PARALLELISM", GdtType.Parallelism, "Parallelism"),
            (@"∥", GdtType.Parallelism, "Parallelism"),
            (@"PERPENDICULARITY", GdtType.Perpendicularity, "Perpendicularity"),
            (@"⊥", GdtType.Perpendicularity, "Perpendicularity"),
            (@"ANGULARITY", GdtType.Angularity, "Angularity"),
            (@"∠", GdtType.Angularity, "Angularity"),

            // Location tolerances
            (@"CONCENTRICITY", GdtType.Concentricity, "Concentricity"),
            (@"◎", GdtType.Concentricity, "Concentricity"),
            (@"SYMMETRY", GdtType.Symmetry, "Symmetry"),
            (@"⌯", GdtType.Symmetry, "Symmetry"),

            // Profile tolerances
            (@"PROFILE\s+OF\s+(?:A\s+)?LINE", GdtType.ProfileOfLine, "Profile of Line"),
            (@"⌒", GdtType.ProfileOfLine, "Profile of Line"),
            (@"PROFILE\s+OF\s+(?:A\s+)?SURFACE", GdtType.ProfileOfSurface, "Profile of Surface"),
            (@"⌓", GdtType.ProfileOfSurface, "Profile of Surface"),
            (@"PROFILE", GdtType.ProfileOfSurface, "Profile"), // Generic "profile" = surface

            // Runout tolerances
            (@"TOTAL\s+RUNOUT", GdtType.TotalRunout, "Total Runout"),
            (@"↗↗", GdtType.TotalRunout, "Total Runout"),
            (@"CIRCULAR\s+RUNOUT", GdtType.CircularRunout, "Circular Runout"),
            (@"RUNOUT", GdtType.CircularRunout, "Runout"),  // Generic "runout" = circular
            (@"↗", GdtType.CircularRunout, "Circular Runout"),
        };

        // Compiled regex for each pattern: TYPE [⌀] TOLERANCE [MMC/LMC] [DATUM A] [B] [C]
        private readonly List<(Regex Regex, GdtType Type, string DisplayName)> _compiledPatterns;

        // Standalone "DATUM A", "DATUM B-C" references
        private static readonly Regex DatumRefPattern = new Regex(
            @"DATUM\s+([A-Z])(?:\s*[-,]\s*([A-Z]))?(?:\s*[-,]\s*([A-Z]))?",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Pattern for tolerance value with optional diameter symbol and material condition
        private const string TolValuePart =
            @"\s*(?:⌀|DIA(?:METER)?\.?\s*)?\.?(\d*\.?\d+)\"?\s*" +
            @"(?:(MMC|LMC|M|L|Ⓜ|Ⓛ)\s*)?" +
            @"(?:(?:TO|W/?R/?T|REF|DATUM)\s+)?(?:([A-Z])\s*(?:(MMC|LMC|M|L|Ⓜ|Ⓛ)\s*)?)?(?:\s*([A-Z])\s*(?:(MMC|LMC|M|L|Ⓜ|Ⓛ)\s*)?)?(?:\s*([A-Z])\s*)?";

        public GdtExtractor()
        {
            _compiledPatterns = new List<(Regex, GdtType, string)>();
            foreach (var (pattern, type, name) in GdtPatterns)
            {
                var regex = new Regex(
                    pattern + TolValuePart,
                    RegexOptions.IgnoreCase | RegexOptions.Compiled);
                _compiledPatterns.Add((regex, type, name));
            }
        }

        /// <summary>
        /// Extracts all GD&amp;T callouts from drawing text.
        /// </summary>
        public List<GdtCallout> Extract(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new List<GdtCallout>();

            var results = new List<GdtCallout>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var (regex, gdtType, displayName) in _compiledPatterns)
            {
                foreach (Match match in regex.Matches(text))
                {
                    if (!TryParseDouble(match.Groups[1].Value, out double tolValue))
                        continue;

                    // Sanity: tolerance should be positive and reasonable
                    if (tolValue <= 0 || tolValue > 1.0) continue;

                    // Dedup by type + value
                    string key = $"{gdtType}:{tolValue:F5}";
                    if (!seen.Add(key)) continue;

                    var callout = new GdtCallout
                    {
                        GdtFeatureType = gdtType,
                        Type = displayName.ToLowerInvariant(),
                        Tolerance = tolValue.ToString("F4", CultureInfo.InvariantCulture),
                        ToleranceValue = tolValue,
                        RawText = match.Value.Trim(),
                        Confidence = 0.80
                    };

                    // Check for diameter symbol
                    callout.IsDiametral = match.Value.Contains("⌀") ||
                        Regex.IsMatch(match.Value, @"DIA(?:METER)?", RegexOptions.IgnoreCase);

                    // Material condition modifiers
                    string mc = match.Groups[2].Success ? match.Groups[2].Value.ToUpper() : "";
                    callout.IsMmc = mc == "MMC" || mc == "M" || mc == "Ⓜ";
                    callout.IsLmc = mc == "LMC" || mc == "L" || mc == "Ⓛ";

                    // Datum references (groups 3, 5, 7)
                    AddDatum(callout, match, 3);
                    AddDatum(callout, match, 5);
                    AddDatum(callout, match, 7);

                    // Classify cost impact
                    ClassifyCallout(callout);

                    results.Add(callout);
                }
            }

            // Sort: tightest tolerances first
            results.Sort((a, b) =>
                (a.ToleranceValue ?? 999).CompareTo(b.ToleranceValue ?? 999));

            return results;
        }

        /// <summary>
        /// Generates cost flags from extracted GD&amp;T callouts.
        /// </summary>
        public List<ToleranceCostFlag> ToCostFlags(List<GdtCallout> callouts)
        {
            if (callouts == null || callouts.Count == 0)
                return new List<ToleranceCostFlag>();

            var flags = new List<ToleranceCostFlag>();

            foreach (var c in callouts)
            {
                if (c.Tier < ToleranceTier.Moderate) continue;

                string datums = c.DatumReferences.Count > 0
                    ? $" (Datum {string.Join("-", c.DatumReferences)})"
                    : "";
                string mmc = c.IsMmc ? " @ MMC" : "";

                flags.Add(new ToleranceCostFlag
                {
                    Description = $"{c.Type} {c.Tolerance}\"{datums}{mmc}",
                    Tier = c.Tier,
                    Impact = c.Impact,
                    SuggestedAction = GetGdtAction(c),
                    Source = "GD&T callout"
                });
            }

            return flags;
        }

        /// <summary>
        /// Generates routing hints from extracted GD&amp;T callouts.
        /// </summary>
        public List<RoutingHint> ToRoutingHints(List<GdtCallout> callouts)
        {
            if (callouts == null || callouts.Count == 0)
                return new List<RoutingHint>();

            var hints = new List<RoutingHint>();
            bool needsCmm = false;
            bool needsFixture = false;
            int tightCount = 0;

            foreach (var c in callouts)
            {
                if (c.Tier >= ToleranceTier.Tight)
                    tightCount++;

                // CMM for any tight geometric tolerance
                if (c.Tier >= ToleranceTier.Tight && !needsCmm)
                {
                    needsCmm = true;
                    hints.Add(new RoutingHint
                    {
                        Operation = RoutingOp.Inspect,
                        WorkCenter = null,
                        NoteText = "CMM INSPECT - GD&T REQUIREMENTS",
                        SourceNote = $"{c.Type} {c.Tolerance}\"",
                        Confidence = 0.90
                    });
                }

                // Fixture for tight true position
                if (c.GdtFeatureType == GdtType.Position && c.Tier >= ToleranceTier.Tight && !needsFixture)
                {
                    needsFixture = true;
                    hints.Add(new RoutingHint
                    {
                        Operation = RoutingOp.Machine,
                        WorkCenter = null,
                        NoteText = "FIXTURE MAY BE REQUIRED - TIGHT TRUE POSITION",
                        SourceNote = $"Position {c.Tolerance}\"",
                        Confidence = 0.75
                    });
                }
            }

            // Flag if multiple tight GD&T callouts (compounds quoting risk)
            if (tightCount >= 3)
            {
                hints.Add(new RoutingHint
                {
                    Operation = RoutingOp.Inspect,
                    WorkCenter = null,
                    NoteText = $"MULTIPLE TIGHT GD&T ({tightCount} CALLOUTS) - REVIEW PRICING",
                    SourceNote = "Multiple callouts",
                    Confidence = 0.85
                });
            }

            return hints;
        }

        // =====================================================================
        // Classification
        // =====================================================================

        private static void ClassifyCallout(GdtCallout callout)
        {
            double tol = callout.ToleranceValue ?? 999;

            // GD&T thresholds are generally tighter expectations than linear dims
            // because they control form/orientation/location
            switch (callout.GdtFeatureType)
            {
                case GdtType.Position:
                    // True position is usually diametral — effective tolerance is tighter
                    callout.Tier = ClassifyPositionTier(tol, callout.IsMmc);
                    break;

                case GdtType.Flatness:
                case GdtType.Straightness:
                case GdtType.Circularity:
                case GdtType.Cylindricity:
                    // Form tolerances — typically tighter thresholds
                    callout.Tier = ClassifyFormTier(tol);
                    break;

                case GdtType.Parallelism:
                case GdtType.Perpendicularity:
                case GdtType.Angularity:
                    // Orientation tolerances
                    callout.Tier = ClassifyOrientationTier(tol);
                    break;

                case GdtType.ProfileOfLine:
                case GdtType.ProfileOfSurface:
                    // Profile — often the tightest requirement on a part
                    callout.Tier = ClassifyProfileTier(tol);
                    break;

                case GdtType.Concentricity:
                case GdtType.Symmetry:
                case GdtType.CircularRunout:
                case GdtType.TotalRunout:
                    // Runout/concentricity — moderate thresholds
                    callout.Tier = ClassifyRunoutTier(tol);
                    break;

                default:
                    callout.Tier = ToleranceTier.Standard;
                    break;
            }

            // Map tier to cost impact
            switch (callout.Tier)
            {
                case ToleranceTier.Precision:
                    callout.Impact = CostImpact.Critical;
                    break;
                case ToleranceTier.Tight:
                    callout.Impact = CostImpact.High;
                    break;
                case ToleranceTier.Moderate:
                    callout.Impact = CostImpact.Medium;
                    break;
                default:
                    callout.Impact = CostImpact.Low;
                    break;
            }
        }

        /// <summary>True position thresholds. MMC gives bonus tolerance so slightly relaxed.</summary>
        public static ToleranceTier ClassifyPositionTier(double tol, bool isMmc)
        {
            // MMC provides bonus tolerance from feature departure from MMC
            // so the effective requirement is less severe
            double effective = isMmc ? tol * 1.5 : tol;

            if (effective <= 0.003) return ToleranceTier.Precision;
            if (effective <= 0.007) return ToleranceTier.Tight;
            if (effective <= 0.014) return ToleranceTier.Moderate;
            return ToleranceTier.Standard;
        }

        /// <summary>Form tolerance thresholds (flatness, straightness, circularity, cylindricity).</summary>
        public static ToleranceTier ClassifyFormTier(double tol)
        {
            if (tol <= 0.001) return ToleranceTier.Precision;
            if (tol <= 0.003) return ToleranceTier.Tight;
            if (tol <= 0.005) return ToleranceTier.Moderate;
            return ToleranceTier.Standard;
        }

        /// <summary>Orientation tolerance thresholds (parallelism, perpendicularity, angularity).</summary>
        public static ToleranceTier ClassifyOrientationTier(double tol)
        {
            if (tol <= 0.002) return ToleranceTier.Precision;
            if (tol <= 0.005) return ToleranceTier.Tight;
            if (tol <= 0.010) return ToleranceTier.Moderate;
            return ToleranceTier.Standard;
        }

        /// <summary>Profile tolerance thresholds.</summary>
        public static ToleranceTier ClassifyProfileTier(double tol)
        {
            if (tol <= 0.002) return ToleranceTier.Precision;
            if (tol <= 0.005) return ToleranceTier.Tight;
            if (tol <= 0.010) return ToleranceTier.Moderate;
            return ToleranceTier.Standard;
        }

        /// <summary>Runout/concentricity thresholds.</summary>
        public static ToleranceTier ClassifyRunoutTier(double tol)
        {
            if (tol <= 0.002) return ToleranceTier.Precision;
            if (tol <= 0.005) return ToleranceTier.Tight;
            if (tol <= 0.010) return ToleranceTier.Moderate;
            return ToleranceTier.Standard;
        }

        private static string GetGdtAction(GdtCallout callout)
        {
            switch (callout.GdtFeatureType)
            {
                case GdtType.Position:
                    if (callout.Tier >= ToleranceTier.Precision)
                        return "CMM INSPECT + FIXTURE REQUIRED";
                    if (callout.Tier >= ToleranceTier.Tight)
                        return "CMM INSPECT REQUIRED";
                    return "VERIFY TRUE POSITION";

                case GdtType.Flatness:
                case GdtType.Straightness:
                    if (callout.Tier >= ToleranceTier.Precision)
                        return "GRINDING OR LAPPING REQUIRED";
                    if (callout.Tier >= ToleranceTier.Tight)
                        return "SURFACE GRINDING MAY BE REQUIRED";
                    return "VERIFY FLATNESS/STRAIGHTNESS";

                case GdtType.ProfileOfLine:
                case GdtType.ProfileOfSurface:
                    if (callout.Tier >= ToleranceTier.Tight)
                        return "CMM INSPECT + POSSIBLE FIXTURE";
                    return "VERIFY PROFILE";

                case GdtType.Concentricity:
                case GdtType.CircularRunout:
                case GdtType.TotalRunout:
                    if (callout.Tier >= ToleranceTier.Tight)
                        return "CMM INSPECT - RUNOUT/CONCENTRICITY";
                    return "VERIFY RUNOUT";

                case GdtType.Perpendicularity:
                case GdtType.Parallelism:
                case GdtType.Angularity:
                    if (callout.Tier >= ToleranceTier.Tight)
                        return "CMM INSPECT - ORIENTATION";
                    return "VERIFY ORIENTATION";

                default:
                    return "REVIEW GD&T REQUIREMENT";
            }
        }

        private static void AddDatum(GdtCallout callout, Match match, int groupIndex)
        {
            if (groupIndex < match.Groups.Count && match.Groups[groupIndex].Success)
            {
                string datum = match.Groups[groupIndex].Value.ToUpper();
                if (datum.Length == 1 && char.IsLetter(datum[0]) && !callout.DatumReferences.Contains(datum))
                    callout.DatumReferences.Add(datum);
            }
        }

        private static bool TryParseDouble(string s, out double value)
        {
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }
    }
}
