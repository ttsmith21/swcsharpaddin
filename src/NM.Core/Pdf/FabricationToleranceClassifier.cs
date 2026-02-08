using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using NM.Core.Pdf.Models;
using static NM.Core.Pdf.IsoToleranceStandard;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Fabrication-aware tolerance classifier for a sheet metal / welding shop.
    ///
    /// Key differences from machining tolerance analysis:
    ///   - Standard tolerances are ISO 13920 BF (weldments) or ISO 2768-m (laser parts)
    ///   - ±0.005" is machining territory, not "standard" — big cost increase
    ///   - Press brake tolerance stackup from bend-to-bend references is a real cost driver
    ///   - AE (ISO 13920) = tighter than standard = extra cost
    ///   - Tighter than AE = probably needs machining = significant cost increase
    ///
    /// This classifier runs AFTER the raw ToleranceAnalyzer and GdtExtractor,
    /// replacing/augmenting cost flags with fabrication-appropriate assessments.
    /// </summary>
    public sealed class FabricationToleranceClassifier
    {
        // Shop defaults (configurable)
        private readonly Iso13920Linear _shopLinearClass;
        private readonly Iso13920Geometric _shopGeometricClass;

        // Regex to detect ISO standard references on drawings
        private static readonly Regex Iso13920Ref = new Regex(
            @"ISO\s*13920\s*[-–:]?\s*([A-D])(?:\s*[-/]?\s*([E-H]))?",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex Iso2768Ref = new Regex(
            @"ISO\s*2768\s*[-–:]?\s*([fmcv])(?:\s*[-/]?\s*([HKL]))?",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Press brake bend indicators
        private static readonly Regex BendCallout = new Regex(
            @"BEND|BRAKE|FOLD|↑|↓|FLANGE|BEND\s*(?:RADIUS|ANGLE|R)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Dimension-to-bend references (e.g., "2.500 REF FROM BEND", tolerance stack)
        private static readonly Regex BendRefDim = new Regex(
            @"(?:FROM|TO|BETWEEN)\s+BEND|BEND\s+(?:TO|LINE)|INSIDE\s+(?:OF\s+)?BEND",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Creates a classifier with shop standard defaults.
        /// </summary>
        /// <param name="shopLinear">Shop standard linear class (default: B)</param>
        /// <param name="shopGeometric">Shop standard geometric class (default: F)</param>
        public FabricationToleranceClassifier(
            Iso13920Linear shopLinear = Iso13920Linear.B,
            Iso13920Geometric shopGeometric = Iso13920Geometric.F)
        {
            _shopLinearClass = shopLinear;
            _shopGeometricClass = shopGeometric;
        }

        /// <summary>
        /// Analyzes a drawing's tolerances from the fabrication shop's perspective.
        /// Runs after ToleranceAnalyzer and GdtExtractor have done raw extraction.
        /// </summary>
        public FabricationToleranceResult Classify(
            string drawingText,
            ToleranceAnalysisResult toleranceResult,
            List<GdtCallout> gdtCallouts)
        {
            var result = new FabricationToleranceResult();

            // Step 1: Detect ISO standard references
            DetectIsoStandards(drawingText, result);

            // Step 2: Classify specific dimension tolerances against fab thresholds
            if (toleranceResult != null)
            {
                ClassifyDimensionTolerances(toleranceResult, result);
            }

            // Step 3: Classify GD&T callouts against fab thresholds
            if (gdtCallouts != null && gdtCallouts.Count > 0)
            {
                ClassifyGdtForFab(gdtCallouts, result);
            }

            // Step 4: Analyze press brake / bend stackup
            AnalyzeBendStackup(drawingText, result);

            // Step 5: Determine overall fabrication tier
            result.OverallTier = DetermineOverallTier(result);

            // Step 6: Generate fab-specific cost flags
            GenerateCostFlags(result);

            return result;
        }

        /// <summary>
        /// Generates routing hints from fabrication tolerance analysis.
        /// </summary>
        public List<RoutingHint> ToRoutingHints(FabricationToleranceResult result)
        {
            var hints = new List<RoutingHint>();
            if (result == null) return hints;

            if (result.RequiresMachining)
            {
                hints.Add(new RoutingHint
                {
                    Operation = RoutingOp.Machine,
                    WorkCenter = null,
                    NoteText = "MACHINING REQUIRED - TOLERANCES TIGHTER THAN FAB STANDARD",
                    SourceNote = "Fabrication tolerance analysis",
                    Confidence = 0.85
                });
            }

            if (result.RequiresCmmInspection)
            {
                hints.Add(new RoutingHint
                {
                    Operation = RoutingOp.Inspect,
                    WorkCenter = null,
                    NoteText = "CMM INSPECT - TIGHT TOLERANCES ON FABRICATED PART",
                    SourceNote = "Fabrication tolerance analysis",
                    Confidence = 0.80
                });
            }

            if (result.BendStackupRisk == BendStackupRisk.High)
            {
                hints.Add(new RoutingHint
                {
                    Operation = RoutingOp.Inspect,
                    WorkCenter = null,
                    NoteText = $"PRESS BRAKE STACKUP RISK - {result.BendCount} BENDS WITH INTER-BEND DIMS",
                    SourceNote = "Bend stackup analysis",
                    Confidence = 0.75
                });
            }

            return hints;
        }

        // =====================================================================
        // Detection
        // =====================================================================

        private void DetectIsoStandards(string text, FabricationToleranceResult result)
        {
            if (string.IsNullOrWhiteSpace(text)) return;

            // ISO 13920 (welding)
            var match13920 = Iso13920Ref.Match(text);
            if (match13920.Success)
            {
                string linearStr = match13920.Groups[1].Value.ToUpper();
                result.Iso13920Detected = true;

                if (Enum.TryParse(linearStr, out Iso13920Linear linear))
                    result.DetectedLinearClass = linear;

                if (match13920.Groups[2].Success)
                {
                    string geoStr = match13920.Groups[2].Value.ToUpper();
                    if (Enum.TryParse(geoStr, out Iso13920Geometric geo))
                        result.DetectedGeometricClass = geo;
                }

                // Evaluate vs shop standard
                if (result.DetectedLinearClass.HasValue)
                {
                    result.LinearTighterThanShop =
                        IsTighterLinear(result.DetectedLinearClass.Value, _shopLinearClass);
                }
                if (result.DetectedGeometricClass.HasValue)
                {
                    result.GeometricTighterThanShop =
                        IsTighterGeometric(result.DetectedGeometricClass.Value, _shopGeometricClass);
                }
            }

            // ISO 2768 (general)
            var match2768 = Iso2768Ref.Match(text);
            if (match2768.Success)
            {
                result.Iso2768Detected = true;
                string clsStr = match2768.Groups[1].Value.ToLower();
                switch (clsStr)
                {
                    case "f": result.DetectedIso2768Class = Iso2768Linear.Fine; break;
                    case "m": result.DetectedIso2768Class = Iso2768Linear.Medium; break;
                    case "c": result.DetectedIso2768Class = Iso2768Linear.Coarse; break;
                    case "v": result.DetectedIso2768Class = Iso2768Linear.VeryCoarse; break;
                }
            }
        }

        // =====================================================================
        // Dimension tolerance classification for fab
        // =====================================================================

        private void ClassifyDimensionTolerances(
            ToleranceAnalysisResult tolResult,
            FabricationToleranceResult fabResult)
        {
            foreach (var dim in tolResult.SpecificTolerances)
            {
                // Convert tolerance band from inches to mm
                double bandMm = InchesToMm(dim.TotalBand);
                double nominalMm = InchesToMm(dim.Nominal);

                // What ISO 13920 class would this tolerance require?
                var requiredClass = ClassifyLinearTolerance13920(nominalMm, bandMm);

                var fabTier = ClassifyForFab(dim.TotalBand, requiredClass);
                fabResult.DimensionClassifications.Add(new FabDimensionClassification
                {
                    Nominal = dim.Nominal,
                    ToleranceBand = dim.TotalBand,
                    ToleranceBandMm = bandMm,
                    RequiredIso13920Class = requiredClass,
                    FabTier = fabTier,
                    RawText = dim.RawText
                });

                if (fabTier == FabricationTier.Machining || fabTier == FabricationTier.PrecisionMachining)
                    fabResult.RequiresMachining = true;
                if (fabTier >= FabricationTier.TighterThanStandard)
                    fabResult.RequiresCmmInspection = true;
            }
        }

        private FabricationTier ClassifyForFab(double toleranceBandInches, Iso13920Linear? requiredClass)
        {
            // If the tolerance is tighter than ISO 13920 Class A,
            // it's definitely machining territory
            double halfBandInches = toleranceBandInches / 2.0;

            if (halfBandInches <= 0.005) // ±0.005" = ±0.127mm — machining
                return FabricationTier.PrecisionMachining;

            if (halfBandInches <= 0.010) // ±0.010" = ±0.254mm — tight machining
                return FabricationTier.Machining;

            if (requiredClass.HasValue)
            {
                if (IsTighterLinear(requiredClass.Value, _shopLinearClass))
                {
                    // Requires class A when shop standard is B
                    if (requiredClass.Value == Iso13920Linear.A)
                        return FabricationTier.TighterThanStandard;
                    return FabricationTier.TighterThanStandard;
                }
            }

            // ±0.030" (±0.76mm) for a 1" dimension → roughly Class B territory
            // This is normal fabrication tolerance
            return FabricationTier.ShopStandard;
        }

        // =====================================================================
        // GD&T classification for fab
        // =====================================================================

        private void ClassifyGdtForFab(List<GdtCallout> callouts, FabricationToleranceResult result)
        {
            foreach (var c in callouts)
            {
                if (!c.ToleranceValue.HasValue) continue;

                double tolMm = InchesToMm(c.ToleranceValue.Value);
                FabricationTier tier;

                switch (c.GdtFeatureType)
                {
                    case GdtType.Flatness:
                    case GdtType.Straightness:
                        // Compare against ISO 13920 geometric (shop standard F)
                        // Typical fab part: 300mm → Class F = 1.5mm
                        if (tolMm <= 0.5) tier = FabricationTier.PrecisionMachining;
                        else if (tolMm <= 1.0) tier = FabricationTier.Machining;
                        else if (tolMm <= 1.5) tier = FabricationTier.TighterThanStandard;
                        else tier = FabricationTier.ShopStandard;
                        break;

                    case GdtType.Perpendicularity:
                    case GdtType.Parallelism:
                    case GdtType.Angularity:
                        // Orientation on welded parts — E class is tight
                        if (tolMm <= 0.5) tier = FabricationTier.PrecisionMachining;
                        else if (tolMm <= 1.0) tier = FabricationTier.Machining;
                        else if (tolMm <= 2.0) tier = FabricationTier.TighterThanStandard;
                        else tier = FabricationTier.ShopStandard;
                        break;

                    case GdtType.Position:
                        // True position on fab parts — very expensive if tight
                        if (tolMm <= 0.5) tier = FabricationTier.PrecisionMachining;
                        else if (tolMm <= 1.5) tier = FabricationTier.Machining;
                        else if (tolMm <= 3.0) tier = FabricationTier.TighterThanStandard;
                        else tier = FabricationTier.ShopStandard;
                        break;

                    default:
                        // Runout, concentricity, profile — all tight for fab
                        if (tolMm <= 0.5) tier = FabricationTier.PrecisionMachining;
                        else if (tolMm <= 1.0) tier = FabricationTier.Machining;
                        else if (tolMm <= 2.0) tier = FabricationTier.TighterThanStandard;
                        else tier = FabricationTier.ShopStandard;
                        break;
                }

                result.GdtClassifications.Add(new FabGdtClassification
                {
                    GdtType = c.GdtFeatureType,
                    ToleranceValue = c.ToleranceValue.Value,
                    ToleranceMm = tolMm,
                    FabTier = tier,
                    RawText = c.RawText
                });

                if (tier >= FabricationTier.Machining) result.RequiresMachining = true;
                if (tier >= FabricationTier.TighterThanStandard) result.RequiresCmmInspection = true;
            }
        }

        // =====================================================================
        // Press brake bend stackup analysis
        // =====================================================================

        private void AnalyzeBendStackup(string text, FabricationToleranceResult result)
        {
            if (string.IsNullOrWhiteSpace(text)) return;

            // Count bend callouts
            result.BendCount = BendCallout.Matches(text).Count;

            // Count bend-to-bend dimension references
            result.BendRefDimCount = BendRefDim.Matches(text).Count;

            // Risk assessment
            if (result.BendCount >= 6 && result.BendRefDimCount >= 2)
                result.BendStackupRisk = BendStackupRisk.High;
            else if (result.BendCount >= 4 && result.BendRefDimCount >= 1)
                result.BendStackupRisk = BendStackupRisk.Medium;
            else if (result.BendCount >= 3)
                result.BendStackupRisk = BendStackupRisk.Low;
            else
                result.BendStackupRisk = BendStackupRisk.None;
        }

        // =====================================================================
        // Overall tier + cost flags
        // =====================================================================

        private FabricationTier DetermineOverallTier(FabricationToleranceResult result)
        {
            var tier = FabricationTier.ShopStandard;

            // ISO 13920 tighter than shop standard
            if (result.LinearTighterThanShop || result.GeometricTighterThanShop)
            {
                if (tier < FabricationTier.TighterThanStandard)
                    tier = FabricationTier.TighterThanStandard;
            }

            // Check dimension classifications
            foreach (var d in result.DimensionClassifications)
                if (d.FabTier > tier) tier = d.FabTier;

            // Check GD&T classifications
            foreach (var g in result.GdtClassifications)
                if (g.FabTier > tier) tier = g.FabTier;

            // Bend stackup can bump tier
            if (result.BendStackupRisk == BendStackupRisk.High &&
                tier < FabricationTier.TighterThanStandard)
                tier = FabricationTier.TighterThanStandard;

            return tier;
        }

        private void GenerateCostFlags(FabricationToleranceResult result)
        {
            // ISO standard comparison
            if (result.Iso13920Detected)
            {
                string detectedClass = "";
                if (result.DetectedLinearClass.HasValue)
                    detectedClass += result.DetectedLinearClass.Value.ToString();
                if (result.DetectedGeometricClass.HasValue)
                    detectedClass += result.DetectedGeometricClass.Value.ToString();

                string shopClass = _shopLinearClass.ToString() + _shopGeometricClass.ToString();

                if (result.LinearTighterThanShop || result.GeometricTighterThanShop)
                {
                    result.CostFlags.Add(new ToleranceCostFlag
                    {
                        Description = $"Drawing specifies ISO 13920-{detectedClass}, shop standard is {shopClass}",
                        Tier = ToleranceTier.Tight,
                        Impact = CostImpact.High,
                        SuggestedAction = "TIGHTER THAN SHOP STANDARD - ADDITIONAL LABOR/SETUP REQUIRED",
                        Source = "ISO 13920 comparison"
                    });
                }
                else
                {
                    result.CostFlags.Add(new ToleranceCostFlag
                    {
                        Description = $"Drawing specifies ISO 13920-{detectedClass} (within shop standard {shopClass})",
                        Tier = ToleranceTier.Standard,
                        Impact = CostImpact.None,
                        SuggestedAction = "STANDARD FABRICATION TOLERANCES",
                        Source = "ISO 13920 comparison"
                    });
                }
            }

            // Machining required flags
            var machiningDims = result.DimensionClassifications
                .Where(d => d.FabTier >= FabricationTier.Machining).ToList();
            if (machiningDims.Count > 0)
            {
                string tightest = machiningDims.OrderBy(d => d.ToleranceBand).First().RawText;
                result.CostFlags.Add(new ToleranceCostFlag
                {
                    Description = $"{machiningDims.Count} dimension(s) require machining (tightest: {tightest})",
                    Tier = ToleranceTier.Precision,
                    Impact = CostImpact.Critical,
                    SuggestedAction = "MACHINING OPERATIONS REQUIRED - SIGNIFICANT COST INCREASE",
                    Source = "Fabrication tolerance analysis"
                });
            }

            // Tighter-than-standard (but not machining)
            var tighterDims = result.DimensionClassifications
                .Where(d => d.FabTier == FabricationTier.TighterThanStandard).ToList();
            if (tighterDims.Count > 0)
            {
                result.CostFlags.Add(new ToleranceCostFlag
                {
                    Description = $"{tighterDims.Count} dimension(s) tighter than shop standard (ISO 13920-A territory)",
                    Tier = ToleranceTier.Moderate,
                    Impact = CostImpact.Medium,
                    SuggestedAction = "EXTRA SETUP/LABOR FOR TIGHTER TOLERANCES",
                    Source = "Fabrication tolerance analysis"
                });
            }

            // GD&T requiring machining
            var machiningGdt = result.GdtClassifications
                .Where(g => g.FabTier >= FabricationTier.Machining).ToList();
            if (machiningGdt.Count > 0)
            {
                result.CostFlags.Add(new ToleranceCostFlag
                {
                    Description = $"{machiningGdt.Count} GD&T callout(s) require machining",
                    Tier = ToleranceTier.Precision,
                    Impact = CostImpact.Critical,
                    SuggestedAction = "GD&T BEYOND FAB CAPABILITY - MACHINING REQUIRED",
                    Source = "Fabrication GD&T analysis"
                });
            }

            // Bend stackup
            if (result.BendStackupRisk >= BendStackupRisk.Medium)
            {
                var impact = result.BendStackupRisk == BendStackupRisk.High
                    ? CostImpact.High : CostImpact.Medium;
                result.CostFlags.Add(new ToleranceCostFlag
                {
                    Description = $"Press brake stackup risk: {result.BendCount} bends, " +
                                  $"{result.BendRefDimCount} bend-to-bend references",
                    Tier = result.BendStackupRisk == BendStackupRisk.High
                        ? ToleranceTier.Tight : ToleranceTier.Moderate,
                    Impact = impact,
                    SuggestedAction = result.BendStackupRisk == BendStackupRisk.High
                        ? "HIGH STACKUP RISK - MAY NEED INTERMEDIATE INSPECTION OR FIXTURE"
                        : "MODERATE STACKUP RISK - VERIFY BEND SEQUENCE",
                    Source = "Press brake analysis"
                });
            }
        }
    }

    // =====================================================================
    // Result models
    // =====================================================================

    /// <summary>
    /// Result of fabrication-aware tolerance classification.
    /// </summary>
    public sealed class FabricationToleranceResult
    {
        // ISO standard detection
        public bool Iso13920Detected { get; set; }
        public Iso13920Linear? DetectedLinearClass { get; set; }
        public Iso13920Geometric? DetectedGeometricClass { get; set; }
        public bool LinearTighterThanShop { get; set; }
        public bool GeometricTighterThanShop { get; set; }

        public bool Iso2768Detected { get; set; }
        public Iso2768Linear? DetectedIso2768Class { get; set; }

        // Dimension classifications
        public List<FabDimensionClassification> DimensionClassifications { get; }
            = new List<FabDimensionClassification>();
        public List<FabGdtClassification> GdtClassifications { get; }
            = new List<FabGdtClassification>();

        // Press brake analysis
        public int BendCount { get; set; }
        public int BendRefDimCount { get; set; }
        public BendStackupRisk BendStackupRisk { get; set; }

        // Overall assessment
        public FabricationTier OverallTier { get; set; }
        public bool RequiresMachining { get; set; }
        public bool RequiresCmmInspection { get; set; }
        public List<ToleranceCostFlag> CostFlags { get; } = new List<ToleranceCostFlag>();

        /// <summary>Human-readable summary.</summary>
        public string Summary
        {
            get
            {
                string iso = Iso13920Detected
                    ? $"ISO 13920-{DetectedLinearClass}{DetectedGeometricClass}"
                    : Iso2768Detected
                        ? $"ISO 2768-{DetectedIso2768Class}"
                        : "No ISO standard specified";
                string machining = RequiresMachining ? ", MACHINING REQUIRED" : "";
                string bends = BendStackupRisk > BendStackupRisk.None
                    ? $", Bend stackup: {BendStackupRisk}"
                    : "";
                return $"{iso}, Tier: {OverallTier}{machining}{bends}, " +
                       $"{CostFlags.Count} cost flags";
            }
        }
    }

    /// <summary>Classification of a specific dimension tolerance in fab context.</summary>
    public sealed class FabDimensionClassification
    {
        public double Nominal { get; set; }
        public double ToleranceBand { get; set; }
        public double ToleranceBandMm { get; set; }
        public Iso13920Linear? RequiredIso13920Class { get; set; }
        public FabricationTier FabTier { get; set; }
        public string RawText { get; set; }
    }

    /// <summary>Classification of a GD&amp;T callout in fab context.</summary>
    public sealed class FabGdtClassification
    {
        public GdtType GdtType { get; set; }
        public double ToleranceValue { get; set; }
        public double ToleranceMm { get; set; }
        public FabricationTier FabTier { get; set; }
        public string RawText { get; set; }
    }

    /// <summary>
    /// Fabrication cost tier — calibrated for sheet metal / welding shop.
    /// </summary>
    public enum FabricationTier
    {
        /// <summary>Within shop standard (ISO 13920-BF or looser). No extra cost.</summary>
        ShopStandard = 0,

        /// <summary>Tighter than shop standard (ISO 13920-AE territory). Extra labor/setup.</summary>
        TighterThanStandard = 1,

        /// <summary>Requires machining operations (±0.010" or tighter). Significant cost increase.</summary>
        Machining = 2,

        /// <summary>Precision machining territory (±0.005" or tighter). Major cost / may need outsource.</summary>
        PrecisionMachining = 3
    }

    /// <summary>
    /// Press brake bend tolerance stackup risk level.
    /// </summary>
    public enum BendStackupRisk
    {
        None = 0,
        Low = 1,
        Medium = 2,
        High = 3
    }
}
