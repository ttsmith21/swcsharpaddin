using System.Linq;
using Xunit;
using NM.Core.Pdf;

namespace NM.Core.Tests.Pdf
{
    public class ToleranceAnalyzerTests
    {
        private readonly ToleranceAnalyzer _analyzer = new ToleranceAnalyzer();

        // ============================================================
        // General tolerance parsing
        // ============================================================

        [Fact]
        public void ParseGeneralTolerance_StandardBlock_ParsesAllPlaces()
        {
            string text = "UNLESS OTHERWISE SPECIFIED TOLERANCES: .X ±0.1  .XX ±0.01  .XXX ±0.005  ANGLES ±1°";
            var gt = _analyzer.ParseGeneralTolerance(text);

            Assert.NotNull(gt);
            Assert.Equal(0.1, gt.OnePlace);
            Assert.Equal(0.01, gt.TwoPlace);
            Assert.Equal(0.005, gt.ThreePlace);
            Assert.Equal(1.0, gt.AngularDegrees);
        }

        [Fact]
        public void ParseGeneralTolerance_FourPlace_DetectsPrecision()
        {
            string text = ".XXXX ±0.0005";
            var gt = _analyzer.ParseGeneralTolerance(text);

            Assert.NotNull(gt);
            Assert.Equal(0.0005, gt.FourPlace);
            Assert.Equal(ToleranceTier.Precision, gt.Tier);
        }

        [Fact]
        public void ParseGeneralTolerance_TwoPlaceOnly_Standard()
        {
            string text = ".XX = ±0.01";
            var gt = _analyzer.ParseGeneralTolerance(text);

            Assert.NotNull(gt);
            Assert.Equal(0.01, gt.TwoPlace);
            Assert.Equal(ToleranceTier.Standard, gt.Tier);
        }

        [Fact]
        public void ParseGeneralTolerance_ThreePlace_TightOrModerate()
        {
            string text = ".XXX ±0.003";
            var gt = _analyzer.ParseGeneralTolerance(text);

            Assert.NotNull(gt);
            Assert.Equal(0.003, gt.ThreePlace);
            // ±0.003 → band = 0.006, which is Moderate
            Assert.True(gt.Tier >= ToleranceTier.Moderate);
        }

        [Fact]
        public void ParseGeneralTolerance_FractionalTolerance()
        {
            string text = "FRACTIONAL ±1/64  .XX ±0.01";
            var gt = _analyzer.ParseGeneralTolerance(text);

            Assert.NotNull(gt);
            Assert.Equal("1/64", gt.FractionalText);
            Assert.NotNull(gt.Fractional);
            Assert.True(gt.Fractional.Value < 0.02); // 1/64 ≈ 0.0156
        }

        [Fact]
        public void ParseGeneralTolerance_AngularMinutes()
        {
            string text = "ANGLES ±30'";
            var gt = _analyzer.ParseGeneralTolerance(text);

            Assert.NotNull(gt);
            Assert.Equal(0.5, gt.AngularDegrees);
        }

        [Fact]
        public void ParseGeneralTolerance_NullInput_ReturnsNull()
        {
            Assert.Null(_analyzer.ParseGeneralTolerance(null));
            Assert.Null(_analyzer.ParseGeneralTolerance(""));
        }

        // ============================================================
        // Specific tolerance extraction
        // ============================================================

        [Fact]
        public void Analyze_BilateralTolerance_Extracts()
        {
            string text = "0.500 ±0.002\n1.000 ±0.005";
            var result = _analyzer.Analyze(text, null);

            Assert.Equal(2, result.SpecificTolerances.Count);

            var tight = result.SpecificTolerances.First(t => t.Nominal == 0.5);
            Assert.Equal(0.002, tight.Plus);
            Assert.Equal(0.002, tight.Minus);
            Assert.Equal(0.004, tight.TotalBand);
            Assert.Equal(ToleranceType.Bilateral, tight.Type);
        }

        [Fact]
        public void Analyze_UnilateralTolerance_Extracts()
        {
            string text = "0.750 +0.002/-0.000";
            var result = _analyzer.Analyze(text, null);

            Assert.NotEmpty(result.SpecificTolerances);
            var tol = result.SpecificTolerances[0];
            Assert.Equal(0.002, tol.Plus);
            Assert.Equal(0.0, tol.Minus);
            Assert.Equal(0.002, tol.TotalBand);
            Assert.Equal(ToleranceType.Unilateral, tol.Type);
        }

        [Fact]
        public void Analyze_SplitUnilateral_Extracts()
        {
            string text = "0.500 +.000 -.002";
            var result = _analyzer.Analyze(text, null);

            Assert.NotEmpty(result.SpecificTolerances);
            var tol = result.SpecificTolerances[0];
            Assert.Equal(0.0, tol.Plus);
            Assert.Equal(0.002, tol.Minus);
            Assert.Equal(ToleranceType.Unilateral, tol.Type);
        }

        [Fact]
        public void Analyze_Deduplicates_SameTolerance()
        {
            string text = "0.500 ±0.002  0.500 ±0.002";
            var result = _analyzer.Analyze(text, null);
            Assert.Single(result.SpecificTolerances);
        }

        // ============================================================
        // Surface finish extraction
        // ============================================================

        [Fact]
        public void Analyze_SurfaceFinishRa_Extracts()
        {
            string text = "SURFACE FINISH Ra 32 UNLESS OTHERWISE NOTED";
            var result = _analyzer.Analyze(text, null);

            Assert.NotEmpty(result.SurfaceFinishCallouts);
            Assert.Equal(32, result.SurfaceFinishCallouts[0].Value);
        }

        [Fact]
        public void Analyze_SurfaceFinishRms_Extracts()
        {
            string text = "125 RMS ALL OVER";
            var result = _analyzer.Analyze(text, null);

            Assert.NotEmpty(result.SurfaceFinishCallouts);
            Assert.Equal(125, result.SurfaceFinishCallouts[0].Value);
            Assert.Equal("RMS", result.SurfaceFinishCallouts[0].Unit);
        }

        [Fact]
        public void Analyze_SurfaceFinish_ClassifiesTier()
        {
            string text = "Ra 16";
            var result = _analyzer.Analyze(text, null);
            Assert.NotEmpty(result.SurfaceFinishCallouts);
            Assert.Equal(ToleranceTier.Precision, result.SurfaceFinishCallouts[0].Tier);
        }

        [Fact]
        public void Analyze_SurfaceFinish125_IsStandard()
        {
            string text = "125 Ra";
            var result = _analyzer.Analyze(text, null);
            Assert.NotEmpty(result.SurfaceFinishCallouts);
            Assert.Equal(ToleranceTier.Standard, result.SurfaceFinishCallouts[0].Tier);
        }

        // ============================================================
        // Cost flag generation
        // ============================================================

        [Fact]
        public void Analyze_TightTolerance_GeneratesCostFlag()
        {
            string text = "0.500 ±0.001";
            var result = _analyzer.Analyze(text, null);

            Assert.True(result.HasCostFlags);
            Assert.Contains(result.CostFlags, f => f.Impact >= CostImpact.High);
        }

        [Fact]
        public void Analyze_StandardTolerance_NoCostFlag()
        {
            string text = "0.500 ±0.010";
            var result = _analyzer.Analyze(text, null);

            Assert.Empty(result.CostFlags);
        }

        [Fact]
        public void Analyze_TightSurfaceFinish_GeneratesCostFlag()
        {
            string text = "Ra 16";
            var result = _analyzer.Analyze(text, null);

            Assert.True(result.HasCostFlags);
            Assert.Contains(result.CostFlags, f => f.Source == "Surface finish callout");
        }

        [Fact]
        public void Analyze_TightGeneralTolerance_GeneratesCostFlag()
        {
            var result = _analyzer.Analyze("SOMETHING", ".XXX ±0.001");

            Assert.True(result.HasCostFlags);
            Assert.Contains(result.CostFlags, f => f.Source == "Title block");
        }

        // ============================================================
        // Tier classification
        // ============================================================

        [Theory]
        [InlineData(0.001, ToleranceTier.Precision)]   // ±0.0005
        [InlineData(0.002, ToleranceTier.Precision)]   // ±0.001
        [InlineData(0.004, ToleranceTier.Tight)]       // ±0.002
        [InlineData(0.005, ToleranceTier.Tight)]       // ±0.0025
        [InlineData(0.008, ToleranceTier.Moderate)]    // ±0.004
        [InlineData(0.010, ToleranceTier.Moderate)]    // ±0.005
        [InlineData(0.020, ToleranceTier.Standard)]    // ±0.010
        [InlineData(0.030, ToleranceTier.Standard)]    // ±0.015
        public void ClassifyDimensionTier_CorrectTier(double totalBand, ToleranceTier expected)
        {
            Assert.Equal(expected, ToleranceAnalyzer.ClassifyDimensionTier(totalBand));
        }

        [Theory]
        [InlineData(8, ToleranceTier.Precision)]
        [InlineData(16, ToleranceTier.Precision)]
        [InlineData(32, ToleranceTier.Tight)]
        [InlineData(63, ToleranceTier.Moderate)]
        [InlineData(125, ToleranceTier.Standard)]
        [InlineData(250, ToleranceTier.Standard)]
        public void ClassifySurfaceFinish_CorrectTier(int raValue, ToleranceTier expected)
        {
            Assert.Equal(expected, ToleranceAnalyzer.ClassifySurfaceFinish(raValue));
        }

        // ============================================================
        // Overall analysis
        // ============================================================

        [Fact]
        public void Analyze_OverallTier_ReflectsTightest()
        {
            string text = "0.500 ±0.010\n1.000 ±0.001\nRa 125";
            var result = _analyzer.Analyze(text, null);

            // 0.001 band → Precision dominates
            Assert.Equal(ToleranceTier.Precision, result.OverallTier);
            Assert.Equal(0.002, result.TightestDimensionBand); // ±0.001 → band 0.002
        }

        [Fact]
        public void Analyze_EmptyText_ReturnsCleanResult()
        {
            var result = _analyzer.Analyze(null, null);

            Assert.Null(result.GeneralTolerance);
            Assert.Empty(result.SpecificTolerances);
            Assert.Empty(result.SurfaceFinishCallouts);
            Assert.Empty(result.CostFlags);
            Assert.Equal(ToleranceTier.Standard, result.OverallTier);
            Assert.False(result.HasCostFlags);
        }

        [Fact]
        public void Analyze_SummaryIsReadable()
        {
            string text = "0.500 ±0.002\nRa 32";
            var result = _analyzer.Analyze(text, ".XXX ±0.005");

            Assert.Contains("General:", result.Summary);
            Assert.Contains("dimension callouts", result.Summary);
            Assert.Contains("finish callouts", result.Summary);
            Assert.Contains("cost flags", result.Summary);
        }

        // ============================================================
        // Routing hints
        // ============================================================

        [Fact]
        public void ToRoutingHints_TightDimension_AddsInspect()
        {
            string text = "0.500 ±0.001";
            var result = _analyzer.Analyze(text, null);
            var hints = _analyzer.ToRoutingHints(result);

            Assert.NotEmpty(hints);
            Assert.Contains(hints, h => h.Operation == RoutingOp.Inspect);
            Assert.Contains(hints, h => h.NoteText.Contains("CMM"));
        }

        [Fact]
        public void ToRoutingHints_TightFinish_AddsGrinding()
        {
            string text = "Ra 16";
            var result = _analyzer.Analyze(text, null);
            var hints = _analyzer.ToRoutingHints(result);

            Assert.NotEmpty(hints);
            Assert.Contains(hints, h => h.NoteText.Contains("GRINDING"));
        }

        [Fact]
        public void ToRoutingHints_NoFlags_ReturnsEmpty()
        {
            string text = "0.500 ±0.010";
            var result = _analyzer.Analyze(text, null);
            var hints = _analyzer.ToRoutingHints(result);

            Assert.Empty(hints);
        }

        [Fact]
        public void ToRoutingHints_NullResult_ReturnsEmpty()
        {
            Assert.Empty(_analyzer.ToRoutingHints(null));
        }

        // ============================================================
        // Realistic drawing text
        // ============================================================

        [Fact]
        public void Analyze_RealisticDrawing()
        {
            string fullText = @"
                PART NO: 12345-01
                DESCRIPTION: MOUNTING BRACKET
                MATERIAL: ASTM A36

                2X 0.375 ±0.002 THRU
                0.750 +.001/-.000
                1.500 ±0.005
                4.000 ±0.010

                SURFACE FINISH 125 RMS UNLESS OTHERWISE NOTED
                Ra 32 ON MATING SURFACES

                NOTES:
                1. BREAK ALL EDGES
                2. DEBURR ALL
            ";

            string tolBlock = ".XX ±0.01  .XXX ±0.005  ANGLES ±1°";

            var result = _analyzer.Analyze(fullText, tolBlock);

            // Should find general tolerance
            Assert.NotNull(result.GeneralTolerance);
            Assert.Equal(0.01, result.GeneralTolerance.TwoPlace);
            Assert.Equal(0.005, result.GeneralTolerance.ThreePlace);

            // Should find specific tolerances
            Assert.True(result.SpecificTolerances.Count >= 3);

            // Should find surface finish
            Assert.True(result.SurfaceFinishCallouts.Count >= 1);

            // Should have cost flags for the tight ones
            Assert.True(result.HasCostFlags);

            // Overall should be at least Tight (0.375 ±0.002 → band 0.004)
            Assert.True(result.OverallTier >= ToleranceTier.Tight);
        }

        [Fact]
        public void Analyze_UosFromFullText_WhenNoTitleBlock()
        {
            string text = @"
                UNLESS OTHERWISE SPECIFIED TOLERANCES:
                .XX ±0.01  .XXX ±0.005
                ANGLES ±0.5°
            ";

            var result = _analyzer.Analyze(text, null);

            // Should extract general tolerance from the full text
            Assert.NotNull(result.GeneralTolerance);
        }
    }
}
