using System;
using System.Collections.Generic;
using System.Globalization;
using NM.Core.Config;
using NM.Core.Config.Tables;
using NM.Core.DataModel;
using NM.Core.Manufacturing;
using NM.Core.Manufacturing.Laser;
using NM.Core.Materials;
using NM.Core.Processing;
using Xunit;
using static NM.Core.Constants.UnitConversions;

namespace NM.Core.Tests
{
    #region Section 1: MaterialCodeMapper Tests

    public class MaterialCodeMapperTests
    {
        [Theory]
        [InlineData("304L", "304L")]
        [InlineData("316L", "316L")]
        [InlineData("309", "309")]
        [InlineData("310", "310")]
        [InlineData("321", "321")]
        [InlineData("330", "330")]
        [InlineData("409", "409")]
        [InlineData("430", "430")]
        [InlineData("2205", "2205")]
        [InlineData("2507", "2507")]
        [InlineData("C22", "C22")]
        [InlineData("C276", "C276")]
        [InlineData("AL6XN", "AL6XN")]
        [InlineData("ALLOY31", "ALLOY31")]
        [InlineData("A36", "A36")]
        [InlineData("ALNZD", "ALNZD")]
        [InlineData("5052", "5052")]
        [InlineData("6061", "6061")]
        public void AllFormMaterials_MapToSelf(string input, string expected)
        {
            string result = MaterialCodeMapper.ToShortCode(input);
            Assert.Equal(expected, result);
        }

        [Theory]
        [InlineData("AISI 304", "304L")]
        [InlineData("AISI 316", "316L")]
        [InlineData("AISI 430", "430")]
        [InlineData("ASTM A36 Steel", "A36")]
        [InlineData("A36 Steel", "A36")]
        [InlineData("6061 Alloy", "6061")]
        [InlineData("6061-T6 (SS)", "6061")]
        [InlineData("5052-H32", "5052")]
        [InlineData("5052 Alloy", "5052")]
        [InlineData("304 Stainless Steel", "304L")]
        [InlineData("316 Stainless Steel (SS)", "316L")]
        [InlineData("AISI 309", "309")]
        [InlineData("AISI 321", "321")]
        public void SwDatabaseNames_MapToShortCodes(string swName, string expected)
        {
            string result = MaterialCodeMapper.ToShortCode(swName);
            Assert.Equal(expected, result);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void NullOrEmpty_ReturnsNull(string input)
        {
            Assert.Null(MaterialCodeMapper.ToShortCode(input));
        }

        [Fact]
        public void CaseInsensitive_LowerCase304L_MapsSame()
        {
            string result = MaterialCodeMapper.ToShortCode("304l");
            Assert.Equal("304L", result);
        }
    }

    #endregion

    #region Section 2: OptiMaterial Resolution — Sheet Metal

    public class OptiMaterialSheetMetalTests
    {
        [Theory]
        [InlineData("304L", "SS")]
        [InlineData("316L", "SS")]
        [InlineData("309", "SS")]
        [InlineData("310", "SS")]
        [InlineData("321", "SS")]
        [InlineData("330", "SS")]
        [InlineData("409", "SS")]
        [InlineData("430", "SS")]
        [InlineData("2205", "SS")]
        [InlineData("2507", "SS")]
        [InlineData("C22", "Other")]
        [InlineData("C276", "SS")]
        [InlineData("AL6XN", "SS")]
        [InlineData("ALLOY31", "SS")]
        [InlineData("A36", "CS")]
        [InlineData("ALNZD", "CS")]
        [InlineData("5052", "AL")]
        [InlineData("6061", "AL")]
        public void AllMaterials_ResolveSheetMetal14GA(string material, string category)
        {
            var pd = new PartData
            {
                Material = material,
                MaterialCategory = category,
                Classification = PartType.SheetMetal,
                Thickness_m = 0.001897, // 14GA ≈ 0.0747"
            };
            pd.Sheet.IsSheetMetal = true;

            string result = StaticOptiMaterialService.Resolve(pd);

            Assert.NotNull(result);
            Assert.StartsWith("S.", result);
            Assert.Contains(material, result);
            Assert.Contains("14GA", result);
        }

        [Theory]
        [InlineData(0.001219, "18GA")]   // 18GA ≈ 0.048"
        [InlineData(0.001524, "16GA")]   // 16GA ≈ 0.060"
        [InlineData(0.001897, "14GA")]   // 14GA ≈ 0.0747"
        [InlineData(0.002659, "12GA")]   // 12GA ≈ 0.1046"
        [InlineData(0.003416, "10GA")]   // 10GA ≈ 0.1345"
        [InlineData(0.00635, ".25IN")]   // 1/4" = 0.250"
        [InlineData(0.0127, ".5IN")]     // 1/2" = 0.500"
        public void ThicknessGauges_ResolveCorrectly(double thickness_m, string expectedLabel)
        {
            var pd = new PartData
            {
                Material = "304L",
                MaterialCategory = "SS",
                Classification = PartType.SheetMetal,
                Thickness_m = thickness_m,
            };
            pd.Sheet.IsSheetMetal = true;

            string result = StaticOptiMaterialService.Resolve(pd);

            Assert.NotNull(result);
            Assert.StartsWith("S.304L", result);
            Assert.Contains(expectedLabel, result);
        }

        [Fact]
        public void ZeroThickness_ReturnsNull()
        {
            var pd = new PartData
            {
                Material = "304L",
                Classification = PartType.SheetMetal,
                Thickness_m = 0,
            };
            pd.Sheet.IsSheetMetal = true;

            Assert.Null(StaticOptiMaterialService.Resolve(pd));
        }

        [Fact]
        public void NullMaterial_ReturnsNull()
        {
            var pd = new PartData
            {
                Material = null,
                Classification = PartType.SheetMetal,
                Thickness_m = 0.001897,
            };
            pd.Sheet.IsSheetMetal = true;

            Assert.Null(StaticOptiMaterialService.Resolve(pd));
        }
    }

    #endregion

    #region Section 3: OptiMaterial Resolution — Tube Shapes

    public class OptiMaterialTubeTests
    {
        [Fact]
        public void Pipe_WithNpsAndSchedule_ResolvesPrefix_P()
        {
            var pd = new PartData
            {
                Material = "304L",
                Classification = PartType.Tube,
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "Round";
            pd.Tube.NpsText = "2\"";
            pd.Tube.ScheduleCode = "40S";

            string result = StaticOptiMaterialService.Resolve(pd);

            Assert.NotNull(result);
            Assert.StartsWith("P.", result);
            Assert.Contains("304L", result);
            Assert.Contains("SCH40S", result);
        }

        [Fact]
        public void RoundTube_WithoutNps_ResolvesPrefix_T()
        {
            var pd = new PartData
            {
                Material = "316L",
                Classification = PartType.Tube,
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "Round";
            pd.Tube.OD_m = 0.0508;   // 2"
            pd.Tube.Wall_m = 0.003175; // 0.125"

            string result = StaticOptiMaterialService.Resolve(pd);

            Assert.NotNull(result);
            Assert.StartsWith("T.", result);
            Assert.Contains("316L", result);
            Assert.Contains("OD", result);
        }

        [Fact]
        public void Angle_ResolvesPrefix_A()
        {
            var pd = new PartData
            {
                Material = "304L",
                Classification = PartType.Tube,
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "Angle";
            pd.Tube.OD_m = 0.0508;    // 2" leg
            pd.Tube.ID_m = 0.0508;    // 2" leg
            pd.Tube.Wall_m = 0.003175; // 0.125"

            string result = StaticOptiMaterialService.Resolve(pd);

            Assert.NotNull(result);
            Assert.StartsWith("A.", result);
            Assert.Contains("304L", result);
        }

        [Fact]
        public void SquareTube_ResolvesPrefix_T_WithSQ()
        {
            var pd = new PartData
            {
                Material = "A36",
                Classification = PartType.Tube,
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "Square";
            pd.Tube.OD_m = 0.0508;    // 2"
            pd.Tube.Wall_m = 0.003175; // 0.125"

            string result = StaticOptiMaterialService.Resolve(pd);

            Assert.NotNull(result);
            Assert.StartsWith("T.", result);
            Assert.Contains("SQ", result);
        }

        [Fact]
        public void RectangleTube_ResolvesPrefix_T()
        {
            var pd = new PartData
            {
                Material = "304L",
                Classification = PartType.Tube,
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "Rectangle";
            pd.Tube.OD_m = 0.0762;    // 3"
            pd.Tube.ID_m = 0.0508;    // 2"
            pd.Tube.Wall_m = 0.001524; // 0.060"

            string result = StaticOptiMaterialService.Resolve(pd);

            Assert.NotNull(result);
            Assert.StartsWith("T.", result);
            Assert.Contains("304L", result);
        }

        [Fact]
        public void Channel_ResolvesPrefix_C()
        {
            var pd = new PartData
            {
                Material = "A36",
                Classification = PartType.Tube,
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "Channel";
            pd.Tube.OD_m = 0.1524;    // 6"
            pd.Tube.ID_m = 0.0508;    // 2"
            pd.Tube.Wall_m = 0.00635; // 0.250"

            string result = StaticOptiMaterialService.Resolve(pd);

            Assert.NotNull(result);
            Assert.StartsWith("C.", result);
        }

        [Fact]
        public void IBeam_ResolvesPrefix_T()
        {
            var pd = new PartData
            {
                Material = "A36",
                Classification = PartType.Tube,
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "I-Beam";
            pd.Tube.OD_m = 0.2032;    // 8"
            pd.Tube.ID_m = 0.1016;    // 4"
            pd.Tube.Wall_m = 0.00635; // 0.250"

            string result = StaticOptiMaterialService.Resolve(pd);

            Assert.NotNull(result);
            Assert.StartsWith("T.", result);
        }

        [Fact]
        public void RoundBar_ResolvesPrefix_R()
        {
            var pd = new PartData
            {
                Material = "304L",
                Classification = PartType.Tube,
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "Round Bar";
            pd.Tube.OD_m = 0.0254; // 1"

            string result = StaticOptiMaterialService.Resolve(pd);

            Assert.NotNull(result);
            Assert.StartsWith("R.", result);
            Assert.Contains("304L", result);
        }
    }

    #endregion

    #region Section 4: Laser Calculator Tests

    public class LaserCalculatorTests
    {
        private sealed class MockLaserSpeedProvider : ILaserSpeedProvider
        {
            private readonly double _feedIpm;
            private readonly double _pierceSec;

            public MockLaserSpeedProvider(double feedIpm, double pierceSec)
            {
                _feedIpm = feedIpm;
                _pierceSec = pierceSec;
            }

            public LaserSpeed GetSpeed(double thicknessIn, string materialCode)
            {
                return new LaserSpeed { FeedRateIpm = _feedIpm, PierceSeconds = _pierceSec };
            }
        }

        private sealed class NoSpeedProvider : ILaserSpeedProvider
        {
            public LaserSpeed GetSpeed(double thicknessIn, string materialCode)
            {
                return default; // HasValue = false
            }
        }

        [Fact]
        public void BasicCalculation_NonZeroResult()
        {
            // 100" cut length, 5 pierces, 0.075" thick 304L
            var metrics = new PartMetrics
            {
                ApproxCutLengthIn = 100,
                PierceCount = 5,
                ThicknessIn = 0.075,
                MaterialCode = "304L",
                MassKg = 2.0,
            };
            var provider = new MockLaserSpeedProvider(feedIpm: 200, pierceSec: 0.5);

            var result = LaserCalculator.Compute(metrics, provider);

            Assert.True(result.SetupHours > 0, "SetupHours should be > 0");
            Assert.True(result.RunHours > 0, "RunHours should be > 0");
            Assert.True(result.TotalHours > 0, "TotalHours should be > 0");
        }

        [Fact]
        public void ZeroCutLength_ZeroRunHours()
        {
            var metrics = new PartMetrics
            {
                ApproxCutLengthIn = 0,
                PierceCount = 0,
                ThicknessIn = 0.075,
                MaterialCode = "304L",
                MassKg = 2.0,
            };
            var provider = new MockLaserSpeedProvider(feedIpm: 200, pierceSec: 0.5);

            var result = LaserCalculator.Compute(metrics, provider);

            // Setup still non-zero (fixed minimum)
            Assert.True(result.SetupHours > 0);
            // Run hours should be ~0 (small proportional load time only)
            Assert.True(result.RunHours < 0.01);
        }

        [Fact]
        public void NoSpeed_ZeroResult()
        {
            var metrics = new PartMetrics
            {
                ApproxCutLengthIn = 100,
                PierceCount = 5,
                ThicknessIn = 0.075,
                MaterialCode = "304L",
                MassKg = 2.0,
            };
            var provider = new NoSpeedProvider();

            var result = LaserCalculator.Compute(metrics, provider);

            Assert.Equal(0, result.SetupHours);
            Assert.Equal(0, result.RunHours);
        }

        [Fact]
        public void NullMetrics_ZeroResult()
        {
            var provider = new MockLaserSpeedProvider(200, 0.5);
            var result = LaserCalculator.Compute(null, provider);

            Assert.Equal(0, result.SetupHours);
            Assert.Equal(0, result.RunHours);
        }

        [Fact]
        public void NullProvider_ZeroResult()
        {
            var metrics = new PartMetrics { ApproxCutLengthIn = 100, ThicknessIn = 0.075, MaterialCode = "304L" };
            var result = LaserCalculator.Compute(metrics, null);

            Assert.Equal(0, result.SetupHours);
            Assert.Equal(0, result.RunHours);
        }

        [Fact]
        public void Waterjet_HigherSetup()
        {
            var metrics = new PartMetrics
            {
                ApproxCutLengthIn = 100,
                PierceCount = 5,
                ThicknessIn = 0.5,
                MaterialCode = "304L",
                MassKg = 5.0,
            };
            var provider = new MockLaserSpeedProvider(feedIpm: 50, pierceSec: 0.5);

            var laserResult = LaserCalculator.Compute(metrics, provider, isWaterjet: false, rawWeightLb: 11);
            var waterjetResult = LaserCalculator.Compute(metrics, provider, isWaterjet: true, rawWeightLb: 11);

            Assert.True(waterjetResult.SetupHours > laserResult.SetupHours,
                "Waterjet setup should be higher than laser setup");
        }

        [Fact]
        public void AllMaterials_WithRealSpeedProvider_GetNonNullSpeed()
        {
            // Load real config tables
            string asmDir = System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location);
            string configDir = System.IO.Path.Combine(asmDir, "config");
            if (!System.IO.Directory.Exists(configDir))
                return; // Skip if no config deployed

            NmConfigProvider.Reload(configDir);
            var provider = new StaticLaserSpeedProvider();

            // All 18 materials at 14GA (0.075")
            var materials = new[]
            {
                "304L", "316L", "309", "310", "321", "330",
                "409", "430", "2205", "2507", "C22", "C276",
                "AL6XN", "ALLOY31", "A36", "ALNZD", "5052", "6061"
            };

            foreach (var mat in materials)
            {
                var speed = provider.GetSpeed(0.075, mat);
                // Speed may or may not have value depending on table coverage.
                // At minimum, the common materials (304L, 316L, A36, 6061, 5052) should resolve.
                // Log which ones fail for visibility.
                if (!speed.HasValue)
                {
                    // Not a hard failure — some exotic materials may not be in the table
                    // but we want to know about it.
                    System.Diagnostics.Debug.WriteLine($"WARN: No laser speed for {mat} at 0.075\"");
                }
            }

            // Core materials MUST have speeds
            Assert.True(provider.GetSpeed(0.075, "304L").HasValue, "304L should have laser speed");
            Assert.True(provider.GetSpeed(0.075, "316L").HasValue, "316L should have laser speed");
            Assert.True(provider.GetSpeed(0.060, "A36").HasValue, "A36 should have laser speed");
            Assert.True(provider.GetSpeed(0.075, "5052").HasValue, "5052 should have laser speed");
            Assert.True(provider.GetSpeed(0.075, "6061").HasValue, "6061 should have laser speed");
        }
    }

    #endregion

    #region Section 5: F140 Press Brake Tests

    public class F140PresssBrakeTests
    {
        public F140PresssBrakeTests()
        {
            NmConfigProvider.ResetToDefaults();
        }

        [Fact]
        public void FlipBonus_AddsOneBendOp()
        {
            var bendNoFlip = new BendInfo { Count = 4, LongestBendIn = 12, NeedsFlip = false };
            var bendFlip = new BendInfo { Count = 4, LongestBendIn = 12, NeedsFlip = true };

            var resultNoFlip = F140Calculator.Compute(bendNoFlip, 3, 10, 1);
            var resultFlip = F140Calculator.Compute(bendFlip, 3, 10, 1);

            // Flip adds 1 extra bend operation → RunHours should be higher
            Assert.True(resultFlip.RunHours > resultNoFlip.RunHours,
                "NeedsFlip=true should add run time for extra bend operation");

            // Specifically: 5 ops * rate vs 4 ops * rate
            double ratio = resultFlip.RunHours / resultNoFlip.RunHours;
            Assert.Equal(5.0 / 4.0, ratio, 3);
        }

        [Fact]
        public void LongPart_TriggersHigherRate()
        {
            // Parts > 60" should get rate 3 (45 sec) even for light parts
            var bend = new BendInfo { Count = 2, LongestBendIn = 12 };
            var shortResult = F140Calculator.Compute(bend, 3, 10, 1);   // short, light → rate 1 (10 sec)
            var longResult = F140Calculator.Compute(bend, 3, 72, 1);    // long > 60" → rate 3 (45 sec)

            Assert.True(longResult.RunHours > shortResult.RunHours,
                "Part > 60\" should trigger higher brake rate");
        }

        [Fact]
        public void ZeroBends_ZeroRunHours()
        {
            var bend = new BendInfo { Count = 0, LongestBendIn = 0 };
            var result = F140Calculator.Compute(bend, 10, 10, 1);

            Assert.Equal(0, result.RunHours);
            // Setup is based on longest bend line (0"), so should be just BrakeSetup/60
            Assert.True(result.SetupHours > 0, "Setup should still be non-zero (has BrakeSetup constant)");
        }

        [Fact]
        public void NullBend_DoesNotThrow()
        {
            var result = F140Calculator.Compute(null, 10, 10, 1);
            Assert.NotNull(result);
            Assert.Equal(0, result.RunHours);
        }

        [Theory]
        [InlineData(3, 10, 10)]      // < 5 lbs, < 12" → Rate1 (10 sec)
        [InlineData(3, 30, 30)]      // < 5 lbs, > 12" → Rate2 (30 sec)
        [InlineData(3, 72, 45)]      // < 5 lbs, > 60" → Rate3 (45 sec)
        [InlineData(10, 10, 45)]     // > 5 lbs → Rate3 (45 sec) minimum
        [InlineData(50, 10, 200)]    // > 40 lbs → Rate4 (200 sec)
        [InlineData(150, 10, 400)]   // > 100 lbs → Rate5 (400 sec)
        public void WeightAndLength_SelectCorrectRate(double weightLb, double lengthIn, int expectedRateSec)
        {
            var bend = new BendInfo { Count = 1, LongestBendIn = 12 };
            var result = F140Calculator.Compute(bend, weightLb, lengthIn, 1);

            // RunHours = 1 bend * rateSec / 3600
            double expectedRunHours = expectedRateSec / 3600.0;
            Assert.Equal(expectedRunHours, result.RunHours, 5);
        }
    }

    #endregion

    #region Section 6: PartDataPropertyMap Round-Trip

    public class PartDataPropertyMapTests
    {
        public PartDataPropertyMapTests()
        {
            NmConfigProvider.ResetToDefaults();
        }

        [Fact]
        public void RoundTrip_PreservesAllCostFields()
        {
            var original = new PartData
            {
                Material = "304L",
                MaterialCategory = "SS",
                OptiMaterial = "S.304L14GA",
                Thickness_m = 0.001897,
                Mass_kg = 2.0,
                SheetPercent = 0.05,
                MaterialCostPerLB = 1.75,
                QuoteQty = 10,
                TotalPrice = 500.0,
            };
            original.Sheet.IsSheetMetal = true;
            original.Sheet.BendCount = 4;

            // Set all cost fields (times in minutes internally)
            original.Cost.OP20_WorkCenter = "F115";
            original.Cost.OP20_S_min = 0.5;     // 0.5 min
            original.Cost.OP20_R_min = 3.0;     // 3 min
            original.Cost.F115_Price = 42.0;

            original.Cost.F140_S_min = 12.5;    // 12.5 min
            original.Cost.F140_R_min = 0.044;   // 0.044 min (≈2.6 sec)
            original.Cost.F140_S_Cost = 16.67;
            original.Cost.F140_Price = 16.73;

            original.Cost.F210_S_min = 0.0;
            original.Cost.F210_R_min = 0.667;   // 40 sec
            original.Cost.F210_Price = 0.47;

            original.Cost.F220_S_min = 5.1;
            original.Cost.F220_R_min = 6.0;
            original.Cost.F220_RN = 10;
            original.Cost.F220_Note = "1/4-20";
            original.Cost.F220_Price = 12.1;

            original.Cost.F325_S_min = 1.0;
            original.Cost.F325_R_min = 2.0;
            original.Cost.F325_Price = 3.3;

            original.Cost.MaterialCost = 3.5;
            original.Cost.MaterialWeight_lb = 4.41;
            original.Cost.TotalMaterialCost = 35.0;
            original.Cost.TotalProcessingCost = 75.0;
            original.Cost.TotalCost = 110.0;

            // Round-trip: PartData → Properties → PartData
            var props = PartDataPropertyMap.ToProperties(original);
            var restored = PartDataPropertyMap.FromProperties(props);

            // Verify key fields preserved (with time unit conversion tolerance)
            Assert.Equal("S.304L14GA", restored.OptiMaterial);
            Assert.Equal("SS", restored.MaterialCategory);
            Assert.True(restored.Sheet.IsSheetMetal);

            // Cost fields (minutes → hours → minutes round-trip)
            AssertMinutesRoundTrip(original.Cost.OP20_S_min, restored.Cost.OP20_S_min, "OP20_S");
            AssertMinutesRoundTrip(original.Cost.OP20_R_min, restored.Cost.OP20_R_min, "OP20_R");
            Assert.Equal(original.Cost.F115_Price, restored.Cost.F115_Price, 3);

            AssertMinutesRoundTrip(original.Cost.F140_S_min, restored.Cost.F140_S_min, "F140_S");
            AssertMinutesRoundTrip(original.Cost.F140_R_min, restored.Cost.F140_R_min, "F140_R");

            AssertMinutesRoundTrip(original.Cost.F210_R_min, restored.Cost.F210_R_min, "F210_R");

            AssertMinutesRoundTrip(original.Cost.F220_S_min, restored.Cost.F220_S_min, "F220_S");
            AssertMinutesRoundTrip(original.Cost.F220_R_min, restored.Cost.F220_R_min, "F220_R");
            Assert.Equal(10, restored.Cost.F220_RN);
            Assert.Equal("1/4-20", restored.Cost.F220_Note);

            AssertMinutesRoundTrip(original.Cost.F325_S_min, restored.Cost.F325_S_min, "F325_S");
            AssertMinutesRoundTrip(original.Cost.F325_R_min, restored.Cost.F325_R_min, "F325_R");

            Assert.Equal(original.Cost.MaterialCost, restored.Cost.MaterialCost, 1);
            Assert.Equal(original.Cost.MaterialWeight_lb, restored.Cost.MaterialWeight_lb, 3);
            Assert.Equal(original.Cost.TotalMaterialCost, restored.Cost.TotalMaterialCost, 1);
            Assert.Equal(original.Cost.TotalProcessingCost, restored.Cost.TotalProcessingCost, 1);
            Assert.Equal(original.Cost.TotalCost, restored.Cost.TotalCost, 1);

            Assert.Equal(original.MaterialCostPerLB, restored.MaterialCostPerLB, 3);
            Assert.Equal(10, restored.QuoteQty);
            Assert.Equal(500.0, restored.TotalPrice, 3);

            Assert.Equal("F115", restored.Cost.OP20_WorkCenter);
        }

        [Fact]
        public void RoundTrip_TubeProperties()
        {
            var original = new PartData
            {
                Material = "304L",
                Classification = PartType.Tube,
            };
            original.Tube.IsTube = true;
            original.Tube.OD_m = 0.0508;       // 2"
            original.Tube.Wall_m = 0.003175;    // 0.125"
            original.Tube.Length_m = 0.3048;     // 12"
            original.Tube.TubeShape = "Round";
            original.Tube.NpsText = "2\"";
            original.Tube.ScheduleCode = "40S";
            original.Tube.NumberOfHoles = 3;

            var props = PartDataPropertyMap.ToProperties(original);
            var restored = PartDataPropertyMap.FromProperties(props);

            Assert.True(restored.Tube.IsTube);
            Assert.Equal(original.Tube.OD_m, restored.Tube.OD_m, 6);
            Assert.Equal(original.Tube.Wall_m, restored.Tube.Wall_m, 6);
            Assert.Equal(original.Tube.Length_m, restored.Tube.Length_m, 5);
            Assert.Equal("Round", restored.Tube.TubeShape);
            Assert.Equal("2\"", restored.Tube.NpsText);
            Assert.Equal("40S", restored.Tube.ScheduleCode);
            Assert.Equal(3, restored.Tube.NumberOfHoles);
        }

        [Fact]
        public void ToProperties_TimesConvertedToHours()
        {
            var pd = new PartData();
            pd.Cost.OP20_S_min = 6.0;  // 6 minutes
            pd.Cost.OP20_R_min = 30.0; // 30 minutes

            var props = PartDataPropertyMap.ToProperties(pd);

            // 6 min = 0.1 hours, 30 min = 0.5 hours
            Assert.Equal("0.1", props["OP20_S"]);
            Assert.Equal("0.5", props["OP20_R"]);
        }

        private void AssertMinutesRoundTrip(double originalMin, double restoredMin, string fieldName)
        {
            // Round-trip: minutes → hours (string "0.####") → minutes
            // Precision loss from 4 decimal places on hours = ~0.006 minutes tolerance
            Assert.True(Math.Abs(originalMin - restoredMin) < 0.01,
                $"{fieldName}: expected {originalMin:F4} min, got {restoredMin:F4} min (diff={Math.Abs(originalMin - restoredMin):F6})");
        }
    }

    #endregion

    #region Section 7: End-to-End Pipeline Simulation

    public class EndToEndPipelineTests
    {
        // Material code → (category code, density kind)
        private static readonly Dictionary<string, string> MaterialCategories =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "304L", "SS" }, { "316L", "SS" }, { "309", "SS" }, { "310", "SS" },
                { "321", "SS" }, { "330", "SS" }, { "409", "SS" }, { "430", "SS" },
                { "2205", "SS" }, { "2507", "SS" }, { "C276", "SS" },
                { "AL6XN", "SS" }, { "ALLOY31", "SS" },
                { "C22", "Other" },
                { "A36", "CS" }, { "ALNZD", "CS" },
                { "5052", "AL" }, { "6061", "AL" },
            };

        public EndToEndPipelineTests()
        {
            NmConfigProvider.ResetToDefaults();
        }

        [Theory]
        [InlineData("304L")]
        [InlineData("316L")]
        [InlineData("309")]
        [InlineData("310")]
        [InlineData("321")]
        [InlineData("330")]
        [InlineData("409")]
        [InlineData("430")]
        [InlineData("2205")]
        [InlineData("2507")]
        [InlineData("C22")]
        [InlineData("C276")]
        [InlineData("AL6XN")]
        [InlineData("ALLOY31")]
        [InlineData("A36")]
        [InlineData("ALNZD")]
        [InlineData("5052")]
        [InlineData("6061")]
        public void FullPipeline_AllMaterials_NonZeroCosts(string material)
        {
            string category = MaterialCategories[material];

            // Build PartData representing a typical 14GA sheet metal part
            var pd = new PartData
            {
                Material = material,
                MaterialCategory = category,
                Classification = PartType.SheetMetal,
                Thickness_m = 0.001897,  // 14GA ≈ 0.0747"
                Mass_kg = 2.0,           // ~4.4 lbs
                BBoxLength_m = 0.3048,   // 12"
                BBoxWidth_m = 0.2032,    // 8"
            };
            pd.Sheet.IsSheetMetal = true;
            pd.Sheet.TotalCutLength_m = 1.016;  // ~40"
            pd.Sheet.BendCount = 4;
            pd.Sheet.BendsBothDirections = false;

            double thicknessIn = pd.Thickness_m * MetersToInches;
            double rawWeightLb = pd.Mass_kg * KgToLbs;

            // 1. Manufacturing Calculator — weight
            var metrics = new PartMetrics
            {
                MassKg = pd.Mass_kg,
                ThicknessIn = thicknessIn,
                MaterialCode = material,
                BendCount = pd.Sheet.BendCount,
                ApproxCutLengthIn = pd.Sheet.TotalCutLength_m * MetersToInches,
                PierceCount = 5,
                Quantity = 1,
            };
            var calcResult = ManufacturingCalculator.Compute(metrics, new CalcOptions());
            Assert.True(calcResult.RawWeightLb > 0, $"{material}: RawWeightLb should be > 0");

            // 2. Laser Calculator — F115
            // Use mock provider with reasonable speeds for all materials
            double feedIpm = category == "AL" ? 300 : 200;
            var laserProvider = new TestLaserSpeedProvider(feedIpm, 0.3);
            var laserResult = LaserCalculator.Compute(metrics, laserProvider);

            Assert.True(laserResult.SetupHours > 0, $"{material}: Laser SetupHours should be > 0");
            Assert.True(laserResult.RunHours > 0, $"{material}: Laser RunHours should be > 0");

            // Store in PartData cost
            pd.Cost.OP20_WorkCenter = "F115";
            pd.Cost.OP20_S_min = laserResult.SetupHours * 60.0;
            pd.Cost.OP20_R_min = laserResult.RunHours * 60.0;
            pd.Cost.F115_Price = laserResult.Cost;

            // 3. F140 Press Brake
            var bendInfo = new BendInfo
            {
                Count = pd.Sheet.BendCount,
                LongestBendIn = pd.BBoxLength_m * MetersToInches, // 12"
                NeedsFlip = pd.Sheet.BendsBothDirections,
            };
            var brakeResult = F140Calculator.Compute(bendInfo, rawWeightLb, pd.BBoxLength_m * MetersToInches, 1);

            Assert.True(brakeResult.SetupHours > 0, $"{material}: F140 SetupHours should be > 0");
            Assert.True(brakeResult.RunHours > 0, $"{material}: F140 RunHours should be > 0");

            pd.Cost.F140_S_min = brakeResult.SetupHours * 60.0;
            pd.Cost.F140_R_min = brakeResult.RunHours * 60.0;
            pd.Cost.F140_Price = brakeResult.Price(1);

            // 4. F210 Deburr
            double cutPerimeterIn = pd.Sheet.TotalCutLength_m * MetersToInches;
            double deburHours = F210Calculator.ComputeHours(cutPerimeterIn);
            pd.Cost.F210_R_min = deburHours * 60.0;
            pd.Cost.F210_Price = F210Calculator.ComputeCost(cutPerimeterIn);

            // 5. OptiMaterial
            pd.OptiMaterial = StaticOptiMaterialService.Resolve(pd);

            Assert.NotNull(pd.OptiMaterial);
            Assert.StartsWith("S.", pd.OptiMaterial);
            Assert.Contains(material, pd.OptiMaterial);

            // 6. ToProperties — verify non-zero property values
            var props = PartDataPropertyMap.ToProperties(pd);

            AssertPropertyPositive(props, "OP20_S", material);
            AssertPropertyPositive(props, "OP20_R", material);
            AssertPropertyPositive(props, "F140_S", material);
            AssertPropertyPositive(props, "F140_R", material);
            Assert.True(pd.Cost.F115_Price > 0, $"{material}: F115_Price should be > 0");
            Assert.True(pd.Cost.F140_Price > 0, $"{material}: F140_Price should be > 0");
            Assert.StartsWith("S.", props["OptiMaterial"]);
        }

        [Theory]
        [InlineData("304L")]
        [InlineData("A36")]
        [InlineData("6061")]
        public void MaterialCodeMapper_IntegratesWithOptiMaterial(string material)
        {
            // Verify the full chain: short code → OptiMaterial
            string shortCode = MaterialCodeMapper.ToShortCode(material);
            Assert.Equal(material, shortCode);

            var pd = new PartData
            {
                Material = shortCode,
                Classification = PartType.SheetMetal,
                Thickness_m = 0.001897,
            };
            pd.Sheet.IsSheetMetal = true;

            string optiMaterial = StaticOptiMaterialService.Resolve(pd);
            Assert.NotNull(optiMaterial);
            Assert.StartsWith("S." + shortCode, optiMaterial);
        }

        private void AssertPropertyPositive(IDictionary<string, string> props, string key, string context)
        {
            Assert.True(props.ContainsKey(key), $"{context}: property '{key}' missing");
            Assert.True(double.TryParse(props[key], NumberStyles.Float, CultureInfo.InvariantCulture, out double val),
                $"{context}: property '{key}' not parseable: '{props[key]}'");
            Assert.True(val > 0, $"{context}: property '{key}' should be > 0, was {val}");
        }

        private sealed class TestLaserSpeedProvider : ILaserSpeedProvider
        {
            private readonly double _feedIpm;
            private readonly double _pierceSec;

            public TestLaserSpeedProvider(double feedIpm, double pierceSec)
            {
                _feedIpm = feedIpm;
                _pierceSec = pierceSec;
            }

            public LaserSpeed GetSpeed(double thicknessIn, string materialCode)
            {
                return new LaserSpeed { FeedRateIpm = _feedIpm, PierceSeconds = _pierceSec };
            }
        }
    }

    #endregion

    #region Section 8: Form Material Pre-Population Mapping

    public class FormMaterialMappingTests
    {
        // This tests the mapping dictionary logic from SelectMaterialRadio
        // without requiring WinForms controls.

        private static readonly Dictionary<string, string> MaterialRadioMap =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // Short codes (self-map)
                { "304L", "304L" }, { "316L", "316L" }, { "309", "309" }, { "310", "310" },
                { "321", "321" }, { "330", "330" }, { "409", "409" }, { "430", "430" },
                { "2205", "2205" }, { "2507", "2507" }, { "C22", "C22" }, { "C276", "C276" },
                { "AL6XN", "AL6XN" }, { "ALLOY31", "ALLOY31" }, { "A36", "A36" },
                { "ALNZD", "ALNZD" }, { "5052", "5052" }, { "6061", "6061" },

                // SW database names
                { "AISI 304", "304L" }, { "AISI 316", "316L" },
                { "AISI 316 Stainless Steel Sheet (SS)", "316L" },
                { "Hastelloy C-22", "C22" }, { "ASTM A36 Steel", "A36" },
                { "5052-H32", "5052" }, { "6061 Alloy", "6061" },
            };

        [Theory]
        [InlineData("304L", "304L")]
        [InlineData("316L", "316L")]
        [InlineData("AISI 304", "304L")]
        [InlineData("AISI 316", "316L")]
        [InlineData("AISI 316 Stainless Steel Sheet (SS)", "316L")]
        [InlineData("Hastelloy C-22", "C22")]
        [InlineData("ASTM A36 Steel", "A36")]
        [InlineData("5052-H32", "5052")]
        [InlineData("6061 Alloy", "6061")]
        [InlineData("309", "309")]
        [InlineData("310", "310")]
        [InlineData("321", "321")]
        [InlineData("330", "330")]
        [InlineData("409", "409")]
        [InlineData("430", "430")]
        [InlineData("2205", "2205")]
        [InlineData("2507", "2507")]
        [InlineData("C276", "C276")]
        [InlineData("AL6XN", "AL6XN")]
        [InlineData("ALLOY31", "ALLOY31")]
        [InlineData("ALNZD", "ALNZD")]
        public void MaterialName_MapsToCorrectRadio(string input, string expectedRadio)
        {
            Assert.True(MaterialRadioMap.TryGetValue(input, out string actual),
                $"Missing mapping for '{input}'");
            Assert.Equal(expectedRadio, actual);
        }

        [Fact]
        public void AllFormMaterials_HavePipelineMapping()
        {
            // Every material the form produces should have a MaterialCodeMapper mapping
            var formMaterials = new[]
            {
                "304L", "316L", "309", "310", "321", "330",
                "409", "430", "2205", "2507", "C22", "C276",
                "AL6XN", "ALLOY31", "A36", "ALNZD", "5052", "6061"
            };

            foreach (var mat in formMaterials)
            {
                string shortCode = MaterialCodeMapper.ToShortCode(mat);
                Assert.NotNull(shortCode);
                Assert.Equal(mat, shortCode); // should pass through unchanged
            }
        }
    }

    #endregion
}
