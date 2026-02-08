using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using NM.Core.Pdf.Models;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Extracts and analyzes tolerances from drawing text:
    ///   1. General tolerance block ("Unless otherwise specified...") → baseline tier
    ///   2. Specific dimension tolerances (±0.002, +.001/-.000) → tight callout flags
    ///   3. Surface finish callouts (Ra 32, 125 RMS) → finish requirements
    ///   4. Angular tolerances (±0.5°, ±30') → angular tier
    /// Generates cost impact flags and routing suggestions for tight tolerances.
    /// </summary>
    public sealed class ToleranceAnalyzer
    {
        // --- General tolerance patterns (from title block) ---
        // Match patterns like: ".XX ±0.01  .XXX ±0.005  ANGLES ±1°"
        private static readonly Regex DecimalPlaceTol = new Regex(
            @"\.X{1,4}\s*(?:=\s*)?[±]\s*(\d*\.?\d+)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex FractionTol = new Regex(
            @"FRACTIONAL?\s*(?:=\s*)?[±]\s*(\d+/\d+)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex AngularTol = new Regex(
            @"ANGLE?S?\s*(?:=\s*)?[±]\s*(\d+(?:\.\d+)?)\s*[°]",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex AngularMinTol = new Regex(
            @"ANGLE?S?\s*(?:=\s*)?[±]\s*(\d+)\s*(?:'|MIN(?:UTES?)?)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // --- Specific tolerance patterns (on individual dimensions) ---

        // ±0.005 or ± .005
        private static readonly Regex BilateralTol = new Regex(
            @"(\d*\.?\d+)\s*[±]\s*(\d*\.?\d+)",
            RegexOptions.Compiled);

        // +0.002/-0.000 or +.001 -.002
        private static readonly Regex UnilateralTol = new Regex(
            @"(\d*\.?\d+)\s*\+\s*(\d*\.?\d+)\s*[/\-]\s*\-?\s*(\d*\.?\d+)",
            RegexOptions.Compiled);

        // .500 +.000 -.002 (separate lines or spaces)
        private static readonly Regex UnilateralSplit = new Regex(
            @"(\d*\.?\d+)\s+\+\s*(\d*\.?\d+)\s+\-\s*(\d*\.?\d+)",
            RegexOptions.Compiled);

        // Surface finish: Ra 32, 63 Ra, RMS 125, 32 RMS, Ra=16
        private static readonly Regex SurfaceFinishRa = new Regex(
            @"(?:Ra\s*=?\s*(\d+))|(?:(\d+)\s*Ra\b)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex SurfaceFinishRms = new Regex(
            @"(?:RMS\s*=?\s*(\d+))|(?:(\d+)\s*RMS\b)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Surface finish symbol value (number followed by check-mark-like context)
        private static readonly Regex SurfaceFinishSymbol = new Regex(
            @"(?:FINISH|SURFACE)\s*(?:=\s*)?(\d+)\s*(?:µ(?:in)?|MICRO(?:INCH)?|RMS|Ra)?",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Analyzes all tolerances in drawing text and title block.
        /// </summary>
        public ToleranceAnalysisResult Analyze(string fullText, string titleBlockToleranceText)
        {
            var result = new ToleranceAnalysisResult();

            // Step 1: Parse general tolerance block
            if (!string.IsNullOrWhiteSpace(titleBlockToleranceText))
            {
                result.GeneralTolerance = ParseGeneralTolerance(titleBlockToleranceText);
            }

            // If general tolerance wasn't in dedicated field, try scanning full text
            if (result.GeneralTolerance == null && !string.IsNullOrWhiteSpace(fullText))
            {
                string uosBlock = ExtractUnlessOtherwiseSpecified(fullText);
                if (uosBlock != null)
                    result.GeneralTolerance = ParseGeneralTolerance(uosBlock);
            }

            if (string.IsNullOrWhiteSpace(fullText))
                return result;

            // Step 2: Find specific dimension tolerances
            ExtractSpecificTolerances(fullText, result);

            // Step 3: Find surface finish callouts
            ExtractSurfaceFinish(fullText, result);

            // Step 4: Classify cost impact
            ClassifyCostImpact(result);

            return result;
        }

        /// <summary>
        /// Parses the general tolerance block text into structured values.
        /// </summary>
        public GeneralTolerance ParseGeneralTolerance(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            var gt = new GeneralTolerance();

            // Find decimal place tolerances (.X, .XX, .XXX, .XXXX)
            var decMatches = DecimalPlaceTol.Matches(text);
            foreach (Match match in decMatches)
            {
                if (TryParseDouble(match.Groups[1].Value, out double tol))
                {
                    int xCount = CountXs(match.Value);
                    switch (xCount)
                    {
                        case 1: gt.OnePlace = tol; break;
                        case 2: gt.TwoPlace = tol; break;
                        case 3: gt.ThreePlace = tol; break;
                        case 4: gt.FourPlace = tol; break;
                    }
                }
            }

            // Fractional tolerance
            var fracMatch = FractionTol.Match(text);
            if (fracMatch.Success)
            {
                gt.FractionalText = fracMatch.Groups[1].Value;
                gt.Fractional = ParseFraction(fracMatch.Groups[1].Value);
            }

            // Angular tolerance in degrees
            var angMatch = AngularTol.Match(text);
            if (angMatch.Success && TryParseDouble(angMatch.Groups[1].Value, out double angDeg))
            {
                gt.AngularDegrees = angDeg;
            }

            // Angular tolerance in minutes
            var angMinMatch = AngularMinTol.Match(text);
            if (angMinMatch.Success && TryParseDouble(angMinMatch.Groups[1].Value, out double angMin))
            {
                gt.AngularDegrees = angMin / 60.0;
            }

            gt.RawText = text.Trim();

            // Determine overall tier from tightest tolerance
            gt.Tier = ClassifyGeneralTier(gt);

            return gt;
        }

        private void ExtractSpecificTolerances(string text, ToleranceAnalysisResult result)
        {
            var seen = new HashSet<string>();

            // Bilateral: 0.500 ±0.002
            foreach (Match match in BilateralTol.Matches(text))
            {
                if (!TryParseDouble(match.Groups[1].Value, out double nominal)) continue;
                if (!TryParseDouble(match.Groups[2].Value, out double tol)) continue;
                if (tol <= 0 || tol >= nominal) continue; // sanity check

                string key = $"B:{nominal:F4}±{tol:F4}";
                if (!seen.Add(key)) continue;

                result.SpecificTolerances.Add(new DimensionTolerance
                {
                    Nominal = nominal,
                    Plus = tol,
                    Minus = tol,
                    TotalBand = tol * 2,
                    Type = ToleranceType.Bilateral,
                    RawText = match.Value.Trim()
                });
            }

            // Unilateral: 0.500 +0.002/-0.000
            foreach (Match match in UnilateralTol.Matches(text))
            {
                if (!TryParseDouble(match.Groups[1].Value, out double nominal)) continue;
                if (!TryParseDouble(match.Groups[2].Value, out double plus)) continue;
                if (!TryParseDouble(match.Groups[3].Value, out double minus)) continue;

                double band = plus + minus;
                if (band <= 0) continue;

                string key = $"U:{nominal:F4}+{plus:F4}-{minus:F4}";
                if (!seen.Add(key)) continue;

                result.SpecificTolerances.Add(new DimensionTolerance
                {
                    Nominal = nominal,
                    Plus = plus,
                    Minus = minus,
                    TotalBand = band,
                    Type = plus == 0 || minus == 0 ? ToleranceType.Unilateral : ToleranceType.Bilateral,
                    RawText = match.Value.Trim()
                });
            }

            // Split unilateral: 0.500 +.000 -.002
            foreach (Match match in UnilateralSplit.Matches(text))
            {
                if (!TryParseDouble(match.Groups[1].Value, out double nominal)) continue;
                if (!TryParseDouble(match.Groups[2].Value, out double plus)) continue;
                if (!TryParseDouble(match.Groups[3].Value, out double minus)) continue;

                double band = plus + minus;
                if (band <= 0) continue;

                string key = $"S:{nominal:F4}+{plus:F4}-{minus:F4}";
                if (!seen.Add(key)) continue;

                result.SpecificTolerances.Add(new DimensionTolerance
                {
                    Nominal = nominal,
                    Plus = plus,
                    Minus = minus,
                    TotalBand = band,
                    Type = ToleranceType.Unilateral,
                    RawText = match.Value.Trim()
                });
            }
        }

        private void ExtractSurfaceFinish(string text, ToleranceAnalysisResult result)
        {
            var seen = new HashSet<int>();

            foreach (var pattern in new[] { SurfaceFinishRa, SurfaceFinishRms, SurfaceFinishSymbol })
            {
                foreach (Match match in pattern.Matches(text))
                {
                    // Try each capture group
                    for (int g = 1; g < match.Groups.Count; g++)
                    {
                        if (!match.Groups[g].Success) continue;
                        if (!int.TryParse(match.Groups[g].Value, out int value)) continue;
                        if (value <= 0 || value > 1000) continue; // sanity
                        if (!seen.Add(value)) continue;

                        result.SurfaceFinishCallouts.Add(new SurfaceFinishCallout
                        {
                            Value = value,
                            Unit = pattern == SurfaceFinishRms ? "RMS" : "Ra",
                            RawText = match.Value.Trim(),
                            Tier = ClassifySurfaceFinish(value)
                        });
                    }
                }
            }
        }

        private void ClassifyCostImpact(ToleranceAnalysisResult result)
        {
            // Flag tight specific tolerances
            foreach (var tol in result.SpecificTolerances)
            {
                tol.Tier = ClassifyDimensionTier(tol.TotalBand);

                if (tol.Tier >= ToleranceTier.Tight)
                {
                    result.CostFlags.Add(new ToleranceCostFlag
                    {
                        Description = $"{tol.RawText} (band: ±{tol.TotalBand / 2:F4}\")",
                        Tier = tol.Tier,
                        Impact = tol.Tier == ToleranceTier.Precision ? CostImpact.Critical : CostImpact.High,
                        SuggestedAction = GetDimensionAction(tol.Tier),
                        Source = "Dimension callout"
                    });
                }
            }

            // Flag tight surface finishes
            foreach (var sf in result.SurfaceFinishCallouts)
            {
                if (sf.Tier >= ToleranceTier.Tight)
                {
                    result.CostFlags.Add(new ToleranceCostFlag
                    {
                        Description = $"Surface finish {sf.Value} {sf.Unit}",
                        Tier = sf.Tier,
                        Impact = sf.Tier == ToleranceTier.Precision ? CostImpact.High : CostImpact.Medium,
                        SuggestedAction = GetSurfaceAction(sf.Tier),
                        Source = "Surface finish callout"
                    });
                }
            }

            // Flag if general tolerance is tight
            if (result.GeneralTolerance != null && result.GeneralTolerance.Tier >= ToleranceTier.Tight)
            {
                result.CostFlags.Add(new ToleranceCostFlag
                {
                    Description = $"General tolerance block: {result.GeneralTolerance.RawText}",
                    Tier = result.GeneralTolerance.Tier,
                    Impact = CostImpact.High,
                    SuggestedAction = "TIGHT GENERAL TOLERANCES - REVIEW ALL DIMENSIONS",
                    Source = "Title block"
                });
            }

            // Overall assessment
            result.TightestDimensionBand = result.SpecificTolerances.Count > 0
                ? result.SpecificTolerances.Min(t => t.TotalBand)
                : (double?)null;
            result.TightestSurfaceFinish = result.SurfaceFinishCallouts.Count > 0
                ? result.SurfaceFinishCallouts.Min(sf => sf.Value)
                : (int?)null;
            result.OverallTier = DetermineOverallTier(result);
        }

        /// <summary>
        /// Generates routing hints from tolerance cost flags.
        /// </summary>
        public List<RoutingHint> ToRoutingHints(ToleranceAnalysisResult result)
        {
            if (result == null || result.CostFlags.Count == 0)
                return new List<RoutingHint>();

            var hints = new List<RoutingHint>();
            bool needsInspect = false;
            bool needsGrinding = false;

            foreach (var flag in result.CostFlags)
            {
                if (flag.Impact >= CostImpact.High && !needsInspect)
                {
                    needsInspect = true;
                    hints.Add(new RoutingHint
                    {
                        Operation = RoutingOp.Inspect,
                        WorkCenter = null,
                        NoteText = "CMM INSPECT - TIGHT TOLERANCES",
                        SourceNote = flag.Description,
                        Confidence = 0.85
                    });
                }

                if (flag.Source == "Surface finish callout" && flag.Tier >= ToleranceTier.Tight && !needsGrinding)
                {
                    needsGrinding = true;
                    hints.Add(new RoutingHint
                    {
                        Operation = RoutingOp.OutsideProcess,
                        WorkCenter = null,
                        NoteText = $"GRINDING REQUIRED - {flag.Description}",
                        SourceNote = flag.Description,
                        Confidence = 0.80
                    });
                }
            }

            return hints;
        }

        // =====================================================================
        // Classification helpers
        // =====================================================================

        private static ToleranceTier ClassifyGeneralTier(GeneralTolerance gt)
        {
            // Use tightest decimal place tolerance
            double? tightest = null;
            if (gt.FourPlace.HasValue) tightest = gt.FourPlace;
            else if (gt.ThreePlace.HasValue) tightest = gt.ThreePlace;
            else if (gt.TwoPlace.HasValue) tightest = gt.TwoPlace;
            else if (gt.OnePlace.HasValue) tightest = gt.OnePlace;

            if (!tightest.HasValue)
                return ToleranceTier.Standard;

            return ClassifyDimensionTier(tightest.Value * 2); // ± value → total band
        }

        public static ToleranceTier ClassifyDimensionTier(double totalBand)
        {
            if (totalBand <= 0.002) return ToleranceTier.Precision;
            if (totalBand <= 0.005) return ToleranceTier.Tight;
            if (totalBand <= 0.010) return ToleranceTier.Moderate;
            return ToleranceTier.Standard;
        }

        public static ToleranceTier ClassifySurfaceFinish(int raValue)
        {
            if (raValue <= 16) return ToleranceTier.Precision;
            if (raValue <= 32) return ToleranceTier.Tight;
            if (raValue <= 63) return ToleranceTier.Moderate;
            return ToleranceTier.Standard;
        }

        private static string GetDimensionAction(ToleranceTier tier)
        {
            switch (tier)
            {
                case ToleranceTier.Precision: return "CMM INSPECT REQUIRED - FIXTURE MAY BE NEEDED";
                case ToleranceTier.Tight: return "CMM INSPECT RECOMMENDED";
                default: return "NOTE ON ROUTING";
            }
        }

        private static string GetSurfaceAction(ToleranceTier tier)
        {
            switch (tier)
            {
                case ToleranceTier.Precision: return "GRINDING OR LAPPING REQUIRED";
                case ToleranceTier.Tight: return "GRINDING MAY BE REQUIRED";
                default: return "VERIFY SURFACE FINISH ACHIEVABLE";
            }
        }

        private static ToleranceTier DetermineOverallTier(ToleranceAnalysisResult result)
        {
            var tier = ToleranceTier.Standard;

            if (result.GeneralTolerance != null && result.GeneralTolerance.Tier > tier)
                tier = result.GeneralTolerance.Tier;

            foreach (var t in result.SpecificTolerances)
                if (t.Tier > tier) tier = t.Tier;

            foreach (var sf in result.SurfaceFinishCallouts)
                if (sf.Tier > tier) tier = sf.Tier;

            return tier;
        }

        private static string ExtractUnlessOtherwiseSpecified(string text)
        {
            var match = Regex.Match(text,
                @"UNLESS\s+OTHERWISE\s+(?:NOTED|SPECIFIED|STATED).{0,500}",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            return match.Success ? match.Value : null;
        }

        private static int CountXs(string text)
        {
            int count = 0;
            bool inXRun = false;
            foreach (char c in text)
            {
                if (c == 'X' || c == 'x')
                {
                    if (!inXRun) inXRun = true;
                    count++;
                }
                else if (inXRun)
                {
                    break;
                }
            }
            return count;
        }

        private static bool TryParseDouble(string s, out double value)
        {
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        private static double? ParseFraction(string fraction)
        {
            var parts = fraction.Split('/');
            if (parts.Length == 2 &&
                double.TryParse(parts[0], out double num) &&
                double.TryParse(parts[1], out double den) &&
                den > 0)
            {
                return num / den;
            }
            return null;
        }
    }

    // =====================================================================
    // Models
    // =====================================================================

    /// <summary>
    /// Full result of tolerance analysis on a drawing.
    /// </summary>
    public sealed class ToleranceAnalysisResult
    {
        /// <summary>General tolerance from the title block.</summary>
        public GeneralTolerance GeneralTolerance { get; set; }

        /// <summary>Specific dimension tolerances found in the drawing.</summary>
        public List<DimensionTolerance> SpecificTolerances { get; } = new List<DimensionTolerance>();

        /// <summary>Surface finish callouts found in the drawing.</summary>
        public List<SurfaceFinishCallout> SurfaceFinishCallouts { get; } = new List<SurfaceFinishCallout>();

        /// <summary>Cost impact flags for tight tolerances.</summary>
        public List<ToleranceCostFlag> CostFlags { get; } = new List<ToleranceCostFlag>();

        /// <summary>Tightest total tolerance band found (inches).</summary>
        public double? TightestDimensionBand { get; set; }

        /// <summary>Tightest surface finish value found (Ra microinches).</summary>
        public int? TightestSurfaceFinish { get; set; }

        /// <summary>Overall tolerance tier for this drawing.</summary>
        public ToleranceTier OverallTier { get; set; }

        /// <summary>True if any tolerance requires special attention.</summary>
        public bool HasCostFlags => CostFlags.Count > 0;

        /// <summary>Human-readable summary.</summary>
        public string Summary
        {
            get
            {
                string gen = GeneralTolerance != null
                    ? $"General: {GeneralTolerance.Tier}"
                    : "General: not specified";
                return $"{gen}, {SpecificTolerances.Count} dimension callouts, " +
                       $"{SurfaceFinishCallouts.Count} finish callouts, " +
                       $"{CostFlags.Count} cost flags, Overall: {OverallTier}";
            }
        }
    }

    /// <summary>
    /// Parsed general tolerance block from the title block.
    /// </summary>
    public sealed class GeneralTolerance
    {
        /// <summary>.X tolerance (e.g., ±0.1).</summary>
        public double? OnePlace { get; set; }
        /// <summary>.XX tolerance (e.g., ±0.01).</summary>
        public double? TwoPlace { get; set; }
        /// <summary>.XXX tolerance (e.g., ±0.005).</summary>
        public double? ThreePlace { get; set; }
        /// <summary>.XXXX tolerance (e.g., ±0.0005).</summary>
        public double? FourPlace { get; set; }
        /// <summary>Fractional tolerance text (e.g., "1/64").</summary>
        public string FractionalText { get; set; }
        /// <summary>Fractional tolerance as decimal.</summary>
        public double? Fractional { get; set; }
        /// <summary>Angular tolerance in degrees.</summary>
        public double? AngularDegrees { get; set; }
        /// <summary>Raw text from the tolerance block.</summary>
        public string RawText { get; set; }
        /// <summary>Overall tier classification.</summary>
        public ToleranceTier Tier { get; set; }
    }

    /// <summary>
    /// A specific dimension tolerance callout found in the drawing.
    /// </summary>
    public sealed class DimensionTolerance
    {
        public double Nominal { get; set; }
        public double Plus { get; set; }
        public double Minus { get; set; }
        /// <summary>Total tolerance band (plus + minus).</summary>
        public double TotalBand { get; set; }
        public ToleranceType Type { get; set; }
        public ToleranceTier Tier { get; set; }
        public string RawText { get; set; }
    }

    /// <summary>
    /// A surface finish callout (Ra or RMS value).
    /// </summary>
    public sealed class SurfaceFinishCallout
    {
        public int Value { get; set; }
        public string Unit { get; set; }
        public ToleranceTier Tier { get; set; }
        public string RawText { get; set; }
    }

    /// <summary>
    /// A cost flag raised by tight tolerance analysis.
    /// </summary>
    public sealed class ToleranceCostFlag
    {
        public string Description { get; set; }
        public ToleranceTier Tier { get; set; }
        public CostImpact Impact { get; set; }
        public string SuggestedAction { get; set; }
        public string Source { get; set; }
    }

    public enum ToleranceType
    {
        Bilateral,
        Unilateral
    }

    public enum ToleranceTier
    {
        /// <summary>Standard shop tolerances (±0.010" or looser).</summary>
        Standard = 0,
        /// <summary>Moderate (±0.005" to ±0.010").</summary>
        Moderate = 1,
        /// <summary>Tight (±0.002" to ±0.005") — may need CMM.</summary>
        Tight = 2,
        /// <summary>Precision (< ±0.002") — definitely needs CMM, possibly grinding.</summary>
        Precision = 3
    }

    public enum CostImpact
    {
        None = 0,
        Low = 1,
        Medium = 2,
        High = 3,
        Critical = 4
    }
}
