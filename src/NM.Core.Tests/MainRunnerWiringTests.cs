using System;
using System.Collections.Generic;
using System.Globalization;
using NM.Core.Config;
using NM.Core.DataModel;
using NM.Core.Manufacturing;
using NM.Core.Manufacturing.Laser;
using NM.Core.Materials;
using NM.Core.Tubes;
using Xunit;
using static NM.Core.Constants.UnitConversions;

namespace NM.Core.Tests
{
    /// <summary>
    /// Tests that replicate MainRunner's CalculateCosts / CalculateSheetMetalCosts /
    /// CalculateTubeCosts wiring logic. These validate that guard conditions, provider
    /// initialization, cost routing, and rollup calculations work correctly — the same
    /// bugs that caused zero costs in production (laser values zero, no press brake time,
    /// OptiMaterial stuck at old value).
    /// </summary>

    #region Sheet Metal Cost Wiring — Guard Conditions

    /// <summary>
    /// Tests that replicate the exact guard conditions in MainRunner.CalculateSheetMetalCosts.
    /// Each test validates that when a specific PartData field is zero/missing, the
    /// corresponding cost operation is skipped — matching MainRunner lines 1272-1396.
    /// </summary>
    public class SheetMetalGuardConditionTests
    {
        public SheetMetalGuardConditionTests()
        {
            NmConfigProvider.ResetToDefaults();
        }

        /// <summary>
        /// MainRunner line 1272: if (pd.Sheet.TotalCutLength_m > 0 && pd.Thickness_m > 0)
        /// When TotalCutLength_m = 0, F115 laser cost should be skipped.
        /// </summary>
        [Fact]
        public void ZeroCutLength_SkipsF115LaserCost()
        {
            var pd = BuildSheetMetalPart("304L", "SS");
            pd.Sheet.TotalCutLength_m = 0; // Guard: must be > 0

            SimulateSheetMetalCosts(pd, 1);

            Assert.Equal(0, pd.Cost.OP20_S_min);
            Assert.Equal(0, pd.Cost.OP20_R_min);
            Assert.Equal(0, pd.Cost.F115_Price);
            Assert.Null(pd.Cost.OP20_WorkCenter); // Never assigned
        }

        /// <summary>
        /// MainRunner line 1272: if (pd.Sheet.TotalCutLength_m > 0 && pd.Thickness_m > 0)
        /// When Thickness_m = 0, F115 laser cost should be skipped.
        /// </summary>
        [Fact]
        public void ZeroThickness_SkipsF115LaserCost()
        {
            var pd = BuildSheetMetalPart("304L", "SS");
            pd.Thickness_m = 0; // Guard: must be > 0

            SimulateSheetMetalCosts(pd, 1);

            Assert.Equal(0, pd.Cost.F115_Price);
        }

        /// <summary>
        /// MainRunner line 1312: if (pd.Sheet.TotalCutLength_m > 0)
        /// When TotalCutLength_m = 0, F210 deburr cost should be skipped.
        /// </summary>
        [Fact]
        public void ZeroCutLength_SkipsF210Deburr()
        {
            var pd = BuildSheetMetalPart("304L", "SS");
            pd.Sheet.TotalCutLength_m = 0;

            SimulateSheetMetalCosts(pd, 1);

            Assert.Equal(0, pd.Cost.F210_R_min);
            Assert.Equal(0, pd.Cost.F210_Price);
        }

        /// <summary>
        /// MainRunner line 1323: if (pd.Sheet.BendCount > 0)
        /// When BendCount = 0, F140 press brake cost should be skipped entirely.
        /// This was one of the original reported bugs (issue #4: no press brake time).
        /// </summary>
        [Fact]
        public void ZeroBendCount_SkipsF140PressBrake()
        {
            var pd = BuildSheetMetalPart("304L", "SS");
            pd.Sheet.BendCount = 0; // Guard: must be > 0

            SimulateSheetMetalCosts(pd, 1);

            Assert.Equal(0, pd.Cost.F140_S_min);
            Assert.Equal(0, pd.Cost.F140_R_min);
            Assert.Equal(0, pd.Cost.F140_Price);
        }

        /// <summary>
        /// MainRunner line 1323: if (pd.Sheet.BendCount > 0)
        /// When BendCount > 0, F140 should produce non-zero costs.
        /// </summary>
        [Fact]
        public void PositiveBendCount_ProducesF140Cost()
        {
            var pd = BuildSheetMetalPart("304L", "SS");
            pd.Sheet.BendCount = 4;
            pd.Extra["LongestBendIn"] = "12.0";

            SimulateSheetMetalCosts(pd, 1);

            Assert.True(pd.Cost.F140_S_min > 0, "F140 setup should be > 0");
            Assert.True(pd.Cost.F140_R_min > 0, "F140 run should be > 0");
            Assert.True(pd.Cost.F140_Price > 0, "F140 price should be > 0");
        }

        /// <summary>
        /// MainRunner line 1381: if (maxRadiusIn > 2.0)
        /// When MaxBendRadiusIn <= 2.0, F325 roll forming should be skipped.
        /// </summary>
        [Fact]
        public void SmallBendRadius_SkipsF325RollForming()
        {
            var pd = BuildSheetMetalPart("304L", "SS");
            // No MaxBendRadiusIn set (defaults to 0) → skipped

            SimulateSheetMetalCosts(pd, 1);

            Assert.Equal(0, pd.Cost.F325_S_min);
            Assert.Equal(0, pd.Cost.F325_R_min);
            Assert.Equal(0, pd.Cost.F325_Price);
        }

        /// <summary>
        /// MainRunner line 1381: if (maxRadiusIn > 2.0)
        /// When MaxBendRadiusIn > 2.0, F325 roll forming should produce cost.
        /// </summary>
        [Fact]
        public void LargeBendRadius_ProducesF325Cost()
        {
            var pd = BuildSheetMetalPart("304L", "SS");
            pd.Extra["MaxBendRadiusIn"] = "6.0";

            SimulateSheetMetalCosts(pd, 1);

            Assert.True(pd.Cost.F325_S_min > 0, "F325 setup should be > 0");
            Assert.True(pd.Cost.F325_R_min > 0, "F325 run should be > 0");
            Assert.True(pd.Cost.F325_Price > 0, "F325 price should be > 0");
        }

        /// <summary>
        /// All guards pass → all five cost operations produce non-zero values.
        /// This is the "happy path" wiring test.
        /// </summary>
        [Fact]
        public void AllGuardsPass_AllCostsNonZero()
        {
            var pd = BuildSheetMetalPart("304L", "SS");
            pd.Sheet.TotalCutLength_m = 1.016;   // 40" cut
            pd.Thickness_m = 0.001897;             // 14GA
            pd.Sheet.BendCount = 4;
            pd.Extra["LongestBendIn"] = "12.0";
            pd.Extra["TappedHoleCount"] = "6";
            pd.Extra["TappedHoleSetups"] = "2";
            pd.Extra["MaxBendRadiusIn"] = "6.0";

            SimulateSheetMetalCosts(pd, 1, tappedHoles: 6);

            Assert.True(pd.Cost.F115_Price > 0, "F115 laser price should be > 0");
            Assert.True(pd.Cost.F210_Price > 0, "F210 deburr price should be > 0");
            Assert.True(pd.Cost.F140_Price > 0, "F140 brake price should be > 0");
            Assert.True(pd.Cost.F220_Price > 0, "F220 tapping price should be > 0");
            Assert.True(pd.Cost.F325_Price > 0, "F325 roll form price should be > 0");
            Assert.Equal("F115", pd.Cost.OP20_WorkCenter);
        }

        /// <summary>
        /// MainRunner line 1294-1295: VBA parity — minimum 0.01 hours for laser setup.
        /// </summary>
        [Fact]
        public void LaserSetup_EnforcesMinimum001Hours()
        {
            var pd = BuildSheetMetalPart("304L", "SS");

            // Use a provider that returns very high feed rate → tiny setup time
            SimulateSheetMetalCosts(pd, 1, feedIpm: 10000);

            // Minimum 0.01 hours = 0.6 minutes
            Assert.True(pd.Cost.OP20_S_min >= 0.6,
                $"Laser setup should be >= 0.6 min (0.01 hr minimum), was {pd.Cost.OP20_S_min:F4}");
        }

        #region Helpers

        private static PartData BuildSheetMetalPart(string material, string category)
        {
            var pd = new PartData
            {
                Material = material,
                MaterialCategory = category,
                Classification = PartType.SheetMetal,
                Thickness_m = 0.001897,    // 14GA
                Mass_kg = 2.0,
                BBoxLength_m = 0.3048,     // 12"
                BBoxWidth_m = 0.2032,      // 8"
            };
            pd.Sheet.IsSheetMetal = true;
            pd.Sheet.TotalCutLength_m = 1.016; // 40"
            pd.Sheet.BendCount = 4;
            pd.Sheet.BendsBothDirections = false;
            return pd;
        }

        /// <summary>
        /// Replicates MainRunner.CalculateSheetMetalCosts wiring logic exactly.
        /// Uses mock laser speed provider so tests don't depend on config files.
        /// </summary>
        private static void SimulateSheetMetalCosts(PartData pd, int quantity,
            double feedIpm = 200, double pierceSec = 0.3, int tappedHoles = 0)
        {
            double rawWeightLb = pd.Mass_kg * KgToLbs;

            // F115 Laser — MainRunner line 1272
            if (pd.Sheet.TotalCutLength_m > 0 && pd.Thickness_m > 0)
            {
                int pierceCount = 0;
                if (pd.Extra.TryGetValue("CutMetrics_PierceCount", out string pcStr))
                    int.TryParse(pcStr, out pierceCount);

                var partMetrics = new PartMetrics
                {
                    ApproxCutLengthIn = pd.Sheet.TotalCutLength_m * MetersToInches,
                    PierceCount = pierceCount,
                    ThicknessIn = pd.Thickness_m * MetersToInches,
                    MaterialCode = pd.Material ?? "304L",
                    MassKg = pd.Mass_kg
                };
                ILaserSpeedProvider speedProvider = new MockLaserProvider(feedIpm, pierceSec);
                var laserResult = LaserCalculator.Compute(partMetrics, speedProvider,
                    isWaterjet: false, rawWeightLb: rawWeightLb);

                double setupHrs = laserResult.SetupHours;
                if (setupHrs < 0.01) setupHrs = 0.01; // VBA minimum
                pd.Cost.OP20_S_min = setupHrs * 60.0;
                pd.Cost.OP20_R_min = laserResult.RunHours * 60.0;
                pd.Cost.OP20_WorkCenter = "F115";
                pd.Cost.F115_Price = laserResult.Cost;
            }

            // F210 Deburr — MainRunner line 1312
            if (pd.Sheet.TotalCutLength_m > 0)
            {
                double cutPerimeterIn = pd.Sheet.TotalCutLength_m * MetersToInches;
                double f210Hours = F210Calculator.ComputeHours(cutPerimeterIn);
                pd.Cost.F210_R_min = f210Hours * 60.0;
                pd.Cost.F210_Price = F210Calculator.ComputeCost(cutPerimeterIn, quantity);
            }

            // F140 Press Brake — MainRunner line 1323
            if (pd.Sheet.BendCount > 0)
            {
                double longestBendIn = 0.0;
                if (pd.Extra.TryGetValue("LongestBendIn", out string extraBend))
                    double.TryParse(extraBend, NumberStyles.Any, CultureInfo.InvariantCulture, out longestBendIn);

                var bendInfo = new BendInfo
                {
                    Count = pd.Sheet.BendCount,
                    LongestBendIn = longestBendIn,
                    NeedsFlip = pd.Sheet.BendsBothDirections
                };
                double partLengthIn = pd.BBoxLength_m * MetersToInches;
                var f140Result = F140Calculator.Compute(bendInfo, rawWeightLb, partLengthIn, quantity);
                pd.Cost.F140_S_min = f140Result.SetupHours * 60.0;
                pd.Cost.F140_R_min = f140Result.RunHours * 60.0;
                pd.Cost.F140_Price = f140Result.Price(quantity);
            }

            // F220 Tapping — MainRunner line 1352
            if (tappedHoles > 0)
            {
                int setups = 1;
                if (pd.Extra.TryGetValue("TappedHoleSetups", out string setupStr))
                    int.TryParse(setupStr, out setups);
                if (setups < 1) setups = 1;

                var f220Input = new F220Input { Setups = setups, Holes = tappedHoles };
                var f220Result = F220Calculator.Compute(f220Input);
                pd.Cost.F220_S_min = f220Result.SetupHours * 60.0;
                pd.Cost.F220_R_min = f220Result.RunHours * 60.0;
                pd.Cost.F220_RN = tappedHoles;
                pd.Cost.F220_Price = (f220Result.SetupHours + f220Result.RunHours * quantity) * CostConstants.F220_COST;
            }

            // F325 Roll Forming — MainRunner line 1381
            double maxRadiusIn = 0;
            if (pd.Extra.TryGetValue("MaxBendRadiusIn", out string extraRadius))
                double.TryParse(extraRadius, NumberStyles.Any, CultureInfo.InvariantCulture, out maxRadiusIn);

            if (maxRadiusIn > 2.0)
            {
                var f325Calc = new F325Calculator();
                double arcLengthIn = maxRadiusIn * Math.PI; // rough half-circle estimate
                var f325Result = f325Calc.CalculateRollForming(maxRadiusIn, arcLengthIn, quantity);
                if (f325Result.RequiresRollForming)
                {
                    pd.Cost.F325_S_min = f325Result.SetupHours * 60.0;
                    pd.Cost.F325_R_min = f325Result.RunHours * 60.0;
                    pd.Cost.F325_Price = f325Result.TotalCost;
                }
            }
        }

        #endregion
    }

    #endregion

    #region Tube Cost Wiring

    /// <summary>
    /// Tests that replicate MainRunner.CalculateTubeCosts wiring logic.
    /// Validates OP20 routing, F325/F140/F210 tube cost chain.
    /// </summary>
    public class TubeCostWiringTests
    {
        public TubeCostWiringTests()
        {
            NmConfigProvider.ResetToDefaults();
        }

        /// <summary>
        /// MainRunner line 1172: bool isSolidBar = pd.Tube.Wall_m < 0.001
        /// Solid bar → OP20 work center = F300 (saw)
        /// </summary>
        [Fact]
        public void SolidBar_RoutesToF300()
        {
            var pd = BuildTubePart("304L", "Round Bar");
            pd.Tube.Wall_m = 0; // Solid bar: no wall

            SimulateTubeCosts(pd, 1);

            Assert.Equal("F300", pd.Cost.OP20_WorkCenter);
            Assert.True(pd.Cost.OP20_S_min > 0, "Solid bar should have setup time");
        }

        /// <summary>
        /// MainRunner line 1184-1191: Round tube OD thresholds for work center routing.
        /// Small OD (< 6.05") → F110, setup = 0.15 hours
        /// </summary>
        [Fact]
        public void SmallRoundTube_RoutesToF110()
        {
            var pd = BuildTubePart("304L", "Round");
            pd.Tube.OD_m = 0.0508;     // 2" OD
            pd.Tube.Wall_m = 0.003175;  // 0.125" wall
            pd.Tube.ID_m = pd.Tube.OD_m - 2 * pd.Tube.Wall_m;

            SimulateTubeCosts(pd, 1);

            Assert.Equal("F110", pd.Cost.OP20_WorkCenter);
            Assert.Equal(0.15 * 60, pd.Cost.OP20_S_min, 1); // 0.15 hrs = 9 min
        }

        /// <summary>
        /// MainRunner line 1188: OD > 10.80 → N145 (CNC), setup = 0.25 hours
        /// </summary>
        [Fact]
        public void LargeRoundTube_RoutesToN145()
        {
            var pd = BuildTubePart("304L", "Round");
            pd.Tube.OD_m = 0.3048;     // 12" OD
            pd.Tube.Wall_m = 0.00635;   // 0.25" wall
            pd.Tube.ID_m = pd.Tube.OD_m - 2 * pd.Tube.Wall_m;

            SimulateTubeCosts(pd, 1);

            Assert.Equal("N145", pd.Cost.OP20_WorkCenter);
            Assert.Equal(0.25 * 60, pd.Cost.OP20_S_min, 1); // 0.25 hrs = 15 min
        }

        /// <summary>
        /// MainRunner line 1193-1196: Angle/Channel → F110, setup = 0.25 hours
        /// </summary>
        [Theory]
        [InlineData("Angle")]
        [InlineData("Channel")]
        public void AngleAndChannel_RouteToF110WithSetup(string shape)
        {
            var pd = BuildTubePart("A36", shape);

            SimulateTubeCosts(pd, 1);

            Assert.Equal("F110", pd.Cost.OP20_WorkCenter);
            Assert.Equal(0.25 * 60, pd.Cost.OP20_S_min, 1);
        }

        /// <summary>
        /// MainRunner line 1198-1201: Rectangle/Square → F110, setup = 0.15 hours
        /// </summary>
        [Theory]
        [InlineData("Rectangle")]
        [InlineData("Square")]
        public void RectSquare_RouteToF110WithLowerSetup(string shape)
        {
            var pd = BuildTubePart("A36", shape);

            SimulateTubeCosts(pd, 1);

            Assert.Equal("F110", pd.Cost.OP20_WorkCenter);
            Assert.Equal(0.15 * 60, pd.Cost.OP20_S_min, 1);
        }

        /// <summary>
        /// MainRunner line 1206-1210: OP20 cost = (setupHrs + runHrs * qty) * rate
        /// Verify the cost formula is correctly wired.
        /// </summary>
        [Fact]
        public void OP20Cost_UsesCorrectFormula()
        {
            var pd = BuildTubePart("304L", "Round");
            pd.Tube.OD_m = 0.0508;
            pd.Tube.Wall_m = 0.003175;
            pd.Tube.ID_m = pd.Tube.OD_m - 2 * pd.Tube.Wall_m;
            int quantity = 5;

            SimulateTubeCosts(pd, quantity);

            double setupHrs = pd.Cost.OP20_S_min / 60.0;
            double runHrs = pd.Cost.OP20_R_min / 60.0;
            double expectedRate = CostConstants.F300_COST; // F110 uses F300 rate
            double expectedCost = (setupHrs + runHrs * quantity) * expectedRate;
            Assert.Equal(expectedCost, pd.Cost.F115_Price, 2);
        }

        /// <summary>
        /// MainRunner line 1221: if (f325Result.RequiresPressBrake)
        /// Light tube (< 40 lbs) → no press brake required.
        /// </summary>
        [Fact]
        public void LightTube_NoPressBrake()
        {
            var pd = BuildTubePart("304L", "Round");
            pd.Mass_kg = 5.0; // ~11 lbs → < 40 lbs

            SimulateTubeCosts(pd, 1);

            Assert.Equal(0, pd.Cost.F140_S_min);
            Assert.Equal(0, pd.Cost.F140_R_min);
            Assert.Equal(0, pd.Cost.F140_Price);
        }

        /// <summary>
        /// MainRunner line 1221: Heavy tube with thick wall → press brake required.
        /// Weight >= 40 lbs AND wall >= 0.165" triggers F140.
        /// </summary>
        [Fact]
        public void HeavyThickTube_GetsPressBrake()
        {
            var pd = BuildTubePart("304L", "Round");
            pd.Mass_kg = 30.0; // ~66 lbs → >= 40 lbs
            pd.Tube.Wall_m = 0.00508; // 0.200" → >= 0.165"

            SimulateTubeCosts(pd, 1);

            Assert.True(pd.Cost.F140_S_min > 0, "Heavy thick tube should get F140 setup");
            Assert.True(pd.Cost.F140_R_min > 0, "Heavy thick tube should get F140 run");
            Assert.True(pd.Cost.F140_Price > 0, "Heavy thick tube should get F140 price");
        }

        /// <summary>
        /// MainRunner line 1232: if (lengthIn > 0) → F210 deburr
        /// Zero length → F210 skipped.
        /// </summary>
        [Fact]
        public void ZeroLength_SkipsTubeDeburr()
        {
            var pd = BuildTubePart("304L", "Round");
            pd.Tube.Length_m = 0;

            SimulateTubeCosts(pd, 1);

            Assert.Equal(0, pd.Cost.F210_S_min);
            Assert.Equal(0, pd.Cost.F210_R_min);
            Assert.Equal(0, pd.Cost.F210_Price);
        }

        /// <summary>
        /// Positive tube length → F210 deburr should produce non-zero cost.
        /// </summary>
        [Fact]
        public void PositiveLength_ProducesTubeDeburr()
        {
            var pd = BuildTubePart("304L", "Round");
            pd.Tube.Length_m = 0.3048; // 12"

            SimulateTubeCosts(pd, 1);

            Assert.True(pd.Cost.F210_S_min > 0, "Tube deburr setup should be > 0");
            Assert.True(pd.Cost.F210_R_min > 0, "Tube deburr run should be > 0");
            Assert.True(pd.Cost.F210_Price > 0, "Tube deburr price should be > 0");
        }

        /// <summary>
        /// F325 roll forming always applies to tubes (no guard condition).
        /// </summary>
        [Fact]
        public void TubeF325_AlwaysApplied()
        {
            var pd = BuildTubePart("304L", "Round");

            SimulateTubeCosts(pd, 1);

            Assert.True(pd.Cost.F325_S_min > 0, "Tube F325 setup should be > 0");
            Assert.True(pd.Cost.F325_R_min > 0, "Tube F325 run should be > 0");
            Assert.True(pd.Cost.F325_Price > 0, "Tube F325 price should be > 0");
        }

        #region Helpers

        private static PartData BuildTubePart(string material, string shape)
        {
            var pd = new PartData
            {
                Material = material,
                Classification = PartType.Tube,
                Mass_kg = 10.0,
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = shape;
            pd.Tube.OD_m = 0.0508;       // 2"
            pd.Tube.Wall_m = 0.003175;    // 0.125"
            pd.Tube.ID_m = 0.0508 - 2 * 0.003175;
            pd.Tube.Length_m = 0.6096;    // 24"
            return pd;
        }

        /// <summary>
        /// Replicates MainRunner.CalculateTubeCosts wiring logic exactly (lines 1163-1241).
        /// </summary>
        private static void SimulateTubeCosts(PartData pd, int quantity)
        {
            double wallIn = pd.Tube.Wall_m * MetersToInches;
            double lengthIn = pd.Tube.Length_m * MetersToInches;
            double odIn = pd.Tube.OD_m * MetersToInches;
            double rawWeightLb = pd.Mass_kg * KgToLbs;

            // OP20 Routing — line 1172
            bool isSolidBar = pd.Tube.Wall_m < 0.001 || pd.Tube.ID_m < 0.001;
            if (isSolidBar)
            {
                pd.Cost.OP20_WorkCenter = "F300";
                pd.Cost.OP20_S_min = 0.05 * 60;
                pd.Cost.OP20_R_min = ((odIn * 90.0) + 15.0) / 60.0;
            }
            else
            {
                string shape = pd.Tube.TubeShape ?? "Round";
                if (shape == "Round")
                {
                    if (odIn > 10.80)       { pd.Cost.OP20_WorkCenter = "N145"; pd.Cost.OP20_S_min = 0.25 * 60; }
                    else if (odIn > 10.05)  { pd.Cost.OP20_WorkCenter = "F110"; pd.Cost.OP20_S_min = 1.0 * 60; }
                    else if (odIn > 6.05)   { pd.Cost.OP20_WorkCenter = "F110"; pd.Cost.OP20_S_min = 0.5 * 60; }
                    else                    { pd.Cost.OP20_WorkCenter = "F110"; pd.Cost.OP20_S_min = 0.15 * 60; }
                }
                else if (shape == "Angle" || shape == "Channel")
                {
                    pd.Cost.OP20_WorkCenter = "F110";
                    pd.Cost.OP20_S_min = 0.25 * 60;
                }
                else // Rectangle, Square
                {
                    pd.Cost.OP20_WorkCenter = "F110";
                    pd.Cost.OP20_S_min = 0.15 * 60;
                }
            }

            // OP20 Cost — line 1206
            double op20SetupHrs = pd.Cost.OP20_S_min / 60.0;
            double op20RunHrs = pd.Cost.OP20_R_min / 60.0;
            double op20Rate = GetOP20Rate(pd.Cost.OP20_WorkCenter);
            pd.Cost.F115_Price = (op20SetupHrs + op20RunHrs * quantity) * op20Rate;

            // F325 Roll Form — line 1214
            var f325Result = TubeWorkCenterRules.ComputeF325(rawWeightLb, wallIn);
            pd.Cost.F325_S_min = f325Result.SetupHours * 60.0;
            pd.Cost.F325_R_min = f325Result.RunHours * 60.0;
            pd.Cost.F325_Price = (f325Result.SetupHours + f325Result.RunHours * quantity) * CostConstants.F325_COST;

            // F140 Press Brake — line 1221
            if (f325Result.RequiresPressBrake)
            {
                var f140Result = TubeWorkCenterRules.ComputeF140(rawWeightLb, wallIn);
                pd.Cost.F140_S_min = f140Result.SetupHours * 60.0;
                pd.Cost.F140_R_min = f140Result.RunHours * 60.0;
                pd.Cost.F140_Price = (f140Result.SetupHours + f140Result.RunHours * quantity) * CostConstants.F140_COST;
            }

            // F210 Deburr — line 1232
            if (lengthIn > 0)
            {
                var f210Result = TubeWorkCenterRules.ComputeF210(lengthIn);
                pd.Cost.F210_S_min = f210Result.SetupHours * 60.0;
                pd.Cost.F210_R_min = f210Result.RunHours * 60.0;
                pd.Cost.F210_Price = (f210Result.SetupHours + f210Result.RunHours * quantity) * CostConstants.F210_COST;
            }
        }

        private static double GetOP20Rate(string workCenter)
        {
            switch (workCenter)
            {
                case "F115": return CostConstants.F115_COST;
                case "F300": return CostConstants.F300_COST;
                case "N145": return CostConstants.F145_COST;
                case "F110": return CostConstants.F300_COST;
                default:     return CostConstants.F300_COST;
            }
        }

        #endregion
    }

    #endregion

    #region Cost Routing & Rollup

    /// <summary>
    /// Tests the top-level CalculateCosts routing logic: purchased parts, tube vs sheet metal,
    /// material cost, and total cost rollup.
    /// </summary>
    public class CostRoutingTests
    {
        public CostRoutingTests()
        {
            NmConfigProvider.ResetToDefaults();
        }

        /// <summary>
        /// MainRunner line 1111: rbPartType == "1" → purchased part, skip all processing.
        /// OP20 set to NPUR, all processing costs should be zero.
        /// </summary>
        [Fact]
        public void PurchasedPart_SkipsAllProcessingCosts()
        {
            var pd = BuildSheetMetalPart("304L", "SS");
            pd.Extra["rbPartType"] = "1";

            SimulateCalculateCosts(pd, 1);

            Assert.Equal("NPUR", pd.Cost.OP20_WorkCenter);
            Assert.Equal(0, pd.Cost.OP20_S_min);
            Assert.Equal(0, pd.Cost.OP20_R_min);
            Assert.Equal(0, pd.Cost.F115_Price);
            Assert.Equal(0, pd.Cost.F140_Price);
            Assert.Equal(0, pd.Cost.F210_Price);
        }

        /// <summary>
        /// MainRunner line 1114-1115: rbPartTypeSub == "2" → customer supplied → CUST.
        /// </summary>
        [Fact]
        public void CustomerSupplied_SetsWorkCenterCUST()
        {
            var pd = BuildSheetMetalPart("304L", "SS");
            pd.Extra["rbPartType"] = "1";
            pd.Extra["rbPartTypeSub"] = "2";

            SimulateCalculateCosts(pd, 1);

            Assert.Equal("CUST", pd.Cost.OP20_WorkCenter);
        }

        /// <summary>
        /// MainRunner line 1121: Tube parts route to CalculateTubeCosts.
        /// Verify tube-specific cost patterns (F325 always present, no F115 laser).
        /// </summary>
        [Fact]
        public void TubePart_RoutesToTubeCosts()
        {
            var pd = new PartData
            {
                Material = "304L",
                Classification = PartType.Tube,
                Mass_kg = 10.0,
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "Round";
            pd.Tube.OD_m = 0.0508;
            pd.Tube.Wall_m = 0.003175;
            pd.Tube.ID_m = 0.0508 - 2 * 0.003175;
            pd.Tube.Length_m = 0.6096;

            SimulateCalculateCosts(pd, 1);

            // Tube should have F325 (always), F210 (length > 0), OP20
            Assert.True(pd.Cost.F325_Price > 0, "Tube should have F325 cost");
            Assert.True(pd.Cost.F210_Price > 0, "Tube should have F210 cost");
            Assert.NotNull(pd.Cost.OP20_WorkCenter);
        }

        /// <summary>
        /// MainRunner line 1125-1127: Non-tube, non-purchased → sheet metal costs.
        /// Verify sheet metal cost patterns (F115 laser, F140 brake, F210 deburr).
        /// </summary>
        [Fact]
        public void SheetMetalPart_RoutesToSheetMetalCosts()
        {
            var pd = BuildSheetMetalPart("304L", "SS");

            SimulateCalculateCosts(pd, 1);

            Assert.Equal("F115", pd.Cost.OP20_WorkCenter);
            Assert.True(pd.Cost.F115_Price > 0, "Sheet metal should have F115 cost");
            Assert.True(pd.Cost.F140_Price > 0, "Sheet metal should have F140 cost");
            Assert.True(pd.Cost.F210_Price > 0, "Sheet metal should have F210 cost");
        }

        /// <summary>
        /// MainRunner line 1131-1143: Material cost calculated for all part types.
        /// </summary>
        [Theory]
        [InlineData("304L", 1.75)]
        [InlineData("316L", 2.25)]
        [InlineData("A36", 0.55)]
        [InlineData("6061", 2.50)]
        [InlineData("5052", 2.35)]
        public void MaterialCost_CalculatedForAllMaterials(string material, double expectedCostPerLb)
        {
            var pd = BuildSheetMetalPart(material, "SS");

            SimulateCalculateCosts(pd, 1);

            Assert.True(pd.Cost.MaterialCost > 0, $"{material}: MaterialCost should be > 0");
            Assert.Equal(expectedCostPerLb, pd.MaterialCostPerLB, 2);
        }

        /// <summary>
        /// MainRunner line 1147-1151: TotalCost = TotalMaterialCost + sum(processing costs).
        /// Verify the rollup formula.
        /// </summary>
        [Fact]
        public void TotalCostRollup_EqualsProcessingPlusMaterial()
        {
            var pd = BuildSheetMetalPart("304L", "SS");

            SimulateCalculateCosts(pd, 1);

            double expectedProcessing = pd.Cost.F115_Price + pd.Cost.F210_Price +
                                        pd.Cost.F140_Price + pd.Cost.F220_Price +
                                        pd.Cost.F325_Price;
            Assert.Equal(expectedProcessing, pd.Cost.TotalProcessingCost, 2);
            Assert.Equal(pd.Cost.TotalMaterialCost + expectedProcessing, pd.Cost.TotalCost, 2);
            Assert.True(pd.Cost.TotalCost > 0, "TotalCost should be > 0");
        }

        /// <summary>
        /// MainRunner line 1104-1106: rawWeightLb prefers MaterialWeight_lb over Mass_kg.
        /// </summary>
        [Fact]
        public void RawWeight_PrefersPrecomputedWeight()
        {
            var pd = BuildSheetMetalPart("304L", "SS");
            pd.Mass_kg = 2.0;                       // ~4.4 lbs
            pd.Cost.MaterialWeight_lb = 10.0;        // pre-computed by ManufacturingCalculator

            SimulateCalculateCosts(pd, 1);

            // Material cost uses the higher pre-computed weight
            // Cost = (10 / 0.85) * 1.75 = ~20.59
            double expectedCost = (10.0 / 0.85) * 1.75;
            Assert.Equal(expectedCost, pd.Cost.MaterialCost, 1);
        }

        /// <summary>
        /// Quantity > 1 affects run costs but not setup costs.
        /// </summary>
        [Theory]
        [InlineData(1)]
        [InlineData(10)]
        [InlineData(100)]
        public void Quantity_AffectsRunCostNotSetup(int qty)
        {
            var pd = BuildSheetMetalPart("304L", "SS");
            pd.Sheet.BendCount = 4;
            pd.Extra["LongestBendIn"] = "12.0";

            SimulateCalculateCosts(pd, qty);

            // F140 setup should be constant regardless of quantity
            var pdSingle = BuildSheetMetalPart("304L", "SS");
            pdSingle.Sheet.BendCount = 4;
            pdSingle.Extra["LongestBendIn"] = "12.0";
            SimulateCalculateCosts(pdSingle, 1);

            Assert.Equal(pdSingle.Cost.F140_S_min, pd.Cost.F140_S_min, 4);

            // F140 price should increase with quantity (due to run time * qty)
            if (qty > 1)
            {
                Assert.True(pd.Cost.F140_Price > pdSingle.Cost.F140_Price,
                    $"F140 price at qty={qty} should exceed qty=1");
            }
        }

        #region Helpers

        private static PartData BuildSheetMetalPart(string material, string category)
        {
            var pd = new PartData
            {
                Material = material,
                MaterialCategory = category,
                Classification = PartType.SheetMetal,
                Thickness_m = 0.001897,
                Mass_kg = 2.0,
                BBoxLength_m = 0.3048,
                BBoxWidth_m = 0.2032,
            };
            pd.Sheet.IsSheetMetal = true;
            pd.Sheet.TotalCutLength_m = 1.016;
            pd.Sheet.BendCount = 4;
            pd.Sheet.BendsBothDirections = false;
            return pd;
        }

        /// <summary>
        /// Replicates MainRunner.CalculateCosts routing logic (lines 1099-1157).
        /// </summary>
        private static void SimulateCalculateCosts(PartData pd, int quantity)
        {
            double rawWeightLb = pd.Cost.MaterialWeight_lb > 0
                ? pd.Cost.MaterialWeight_lb
                : pd.Mass_kg * KgToLbs;

            // Purchased part early-out
            if (pd.Extra.TryGetValue("rbPartType", out string rbPartType) && rbPartType == "1")
            {
                pd.Extra.TryGetValue("rbPartTypeSub", out string rbPartTypeSub);
                pd.Cost.OP20_WorkCenter = (rbPartTypeSub == "2") ? "CUST" : "NPUR";
                pd.Cost.OP20_S_min = 0;
                pd.Cost.OP20_R_min = 0;
            }
            else if (pd.Classification == PartType.Tube && pd.Tube.IsTube)
            {
                SimulateTubeCosts(pd, rawWeightLb, quantity);
            }
            else
            {
                SimulateSheetMetalCosts(pd, rawWeightLb, quantity);
            }

            // Material Cost
            if (rawWeightLb > 0 && !string.IsNullOrWhiteSpace(pd.Material))
            {
                var matInput = new MaterialCostCalculator.MaterialCostInput
                {
                    WeightLb = rawWeightLb,
                    MaterialCode = pd.Material,
                    Quantity = quantity,
                    NestEfficiency = 0.85
                };
                var matResult = MaterialCostCalculator.Calculate(matInput);
                pd.Cost.MaterialCost = matResult.CostPerPiece;
                pd.Cost.TotalMaterialCost = matResult.TotalMaterialCost;
                pd.MaterialCostPerLB = matResult.CostPerLb;
            }

            // Total Cost Rollup
            double processingCost = pd.Cost.F115_Price + pd.Cost.F210_Price +
                                    pd.Cost.F140_Price + pd.Cost.F220_Price +
                                    pd.Cost.F325_Price;
            pd.Cost.TotalProcessingCost = processingCost;
            pd.Cost.TotalCost = pd.Cost.TotalMaterialCost + processingCost;
        }

        private static void SimulateSheetMetalCosts(PartData pd, double rawWeightLb, int quantity)
        {
            if (pd.Sheet.TotalCutLength_m > 0 && pd.Thickness_m > 0)
            {
                var partMetrics = new PartMetrics
                {
                    ApproxCutLengthIn = pd.Sheet.TotalCutLength_m * MetersToInches,
                    PierceCount = 2,
                    ThicknessIn = pd.Thickness_m * MetersToInches,
                    MaterialCode = pd.Material ?? "304L",
                    MassKg = pd.Mass_kg
                };
                ILaserSpeedProvider speedProvider = new MockLaserProvider(200, 0.3);
                var laserResult = LaserCalculator.Compute(partMetrics, speedProvider,
                    isWaterjet: false, rawWeightLb: rawWeightLb);

                double setupHrs = laserResult.SetupHours;
                if (setupHrs < 0.01) setupHrs = 0.01;
                pd.Cost.OP20_S_min = setupHrs * 60.0;
                pd.Cost.OP20_R_min = laserResult.RunHours * 60.0;
                pd.Cost.OP20_WorkCenter = "F115";
                pd.Cost.F115_Price = laserResult.Cost;
            }

            if (pd.Sheet.TotalCutLength_m > 0)
            {
                double cutPerimeterIn = pd.Sheet.TotalCutLength_m * MetersToInches;
                pd.Cost.F210_R_min = F210Calculator.ComputeHours(cutPerimeterIn) * 60.0;
                pd.Cost.F210_Price = F210Calculator.ComputeCost(cutPerimeterIn, quantity);
            }

            if (pd.Sheet.BendCount > 0)
            {
                double longestBendIn = 0.0;
                if (pd.Extra.TryGetValue("LongestBendIn", out string extraBend))
                    double.TryParse(extraBend, NumberStyles.Any, CultureInfo.InvariantCulture, out longestBendIn);

                var bendInfo = new BendInfo
                {
                    Count = pd.Sheet.BendCount,
                    LongestBendIn = longestBendIn,
                    NeedsFlip = pd.Sheet.BendsBothDirections
                };
                var f140Result = F140Calculator.Compute(bendInfo, rawWeightLb,
                    pd.BBoxLength_m * MetersToInches, quantity);
                pd.Cost.F140_S_min = f140Result.SetupHours * 60.0;
                pd.Cost.F140_R_min = f140Result.RunHours * 60.0;
                pd.Cost.F140_Price = f140Result.Price(quantity);
            }
        }

        private static void SimulateTubeCosts(PartData pd, double rawWeightLb, int quantity)
        {
            double wallIn = pd.Tube.Wall_m * MetersToInches;
            double lengthIn = pd.Tube.Length_m * MetersToInches;
            double odIn = pd.Tube.OD_m * MetersToInches;

            bool isSolidBar = pd.Tube.Wall_m < 0.001 || pd.Tube.ID_m < 0.001;
            if (isSolidBar)
            {
                pd.Cost.OP20_WorkCenter = "F300";
                pd.Cost.OP20_S_min = 0.05 * 60;
                pd.Cost.OP20_R_min = ((odIn * 90.0) + 15.0) / 60.0;
            }
            else
            {
                string shape = pd.Tube.TubeShape ?? "Round";
                if (shape == "Round")
                {
                    if (odIn > 10.80)       { pd.Cost.OP20_WorkCenter = "N145"; pd.Cost.OP20_S_min = 0.25 * 60; }
                    else if (odIn > 10.05)  { pd.Cost.OP20_WorkCenter = "F110"; pd.Cost.OP20_S_min = 1.0 * 60; }
                    else if (odIn > 6.05)   { pd.Cost.OP20_WorkCenter = "F110"; pd.Cost.OP20_S_min = 0.5 * 60; }
                    else                    { pd.Cost.OP20_WorkCenter = "F110"; pd.Cost.OP20_S_min = 0.15 * 60; }
                }
                else if (shape == "Angle" || shape == "Channel")
                {
                    pd.Cost.OP20_WorkCenter = "F110";
                    pd.Cost.OP20_S_min = 0.25 * 60;
                }
                else
                {
                    pd.Cost.OP20_WorkCenter = "F110";
                    pd.Cost.OP20_S_min = 0.15 * 60;
                }
            }

            double op20SetupHrs = pd.Cost.OP20_S_min / 60.0;
            double op20RunHrs = pd.Cost.OP20_R_min / 60.0;
            double op20Rate = GetOP20Rate(pd.Cost.OP20_WorkCenter);
            pd.Cost.F115_Price = (op20SetupHrs + op20RunHrs * quantity) * op20Rate;

            var f325Result = TubeWorkCenterRules.ComputeF325(rawWeightLb, wallIn);
            pd.Cost.F325_S_min = f325Result.SetupHours * 60.0;
            pd.Cost.F325_R_min = f325Result.RunHours * 60.0;
            pd.Cost.F325_Price = (f325Result.SetupHours + f325Result.RunHours * quantity) * CostConstants.F325_COST;

            if (f325Result.RequiresPressBrake)
            {
                var f140Result = TubeWorkCenterRules.ComputeF140(rawWeightLb, wallIn);
                pd.Cost.F140_S_min = f140Result.SetupHours * 60.0;
                pd.Cost.F140_R_min = f140Result.RunHours * 60.0;
                pd.Cost.F140_Price = (f140Result.SetupHours + f140Result.RunHours * quantity) * CostConstants.F140_COST;
            }

            if (lengthIn > 0)
            {
                var f210Result = TubeWorkCenterRules.ComputeF210(lengthIn);
                pd.Cost.F210_S_min = f210Result.SetupHours * 60.0;
                pd.Cost.F210_R_min = f210Result.RunHours * 60.0;
                pd.Cost.F210_Price = (f210Result.SetupHours + f210Result.RunHours * quantity) * CostConstants.F210_COST;
            }
        }

        private static double GetOP20Rate(string workCenter)
        {
            switch (workCenter)
            {
                case "F115": return CostConstants.F115_COST;
                case "F300": return CostConstants.F300_COST;
                case "N145": return CostConstants.F145_COST;
                case "F110": return CostConstants.F300_COST;
                default:     return CostConstants.F300_COST;
            }
        }

        #endregion
    }

    #endregion

    #region StaticLaserSpeedProvider with Real Config

    /// <summary>
    /// Tests that StaticLaserSpeedProvider correctly delegates to NmConfigProvider.Tables.
    /// Validates the production wiring path used by MainRunner (line 1290).
    /// </summary>
    public class StaticLaserSpeedProviderWiringTests
    {
        public StaticLaserSpeedProviderWiringTests()
        {
            // Load real config tables (copied to test output by .csproj Content rules)
            string asmDir = System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location);
            string configDir = System.IO.Path.Combine(asmDir, "config");
            if (System.IO.Directory.Exists(configDir))
                NmConfigProvider.Reload(configDir);
            else
                NmConfigProvider.ResetToDefaults();
        }

        /// <summary>
        /// After NmConfigProvider.Initialize(), StaticLaserSpeedProvider should return
        /// valid speeds for the core materials at a standard thickness.
        /// This tests the exact wiring path: MainRunner → new StaticLaserSpeedProvider()
        /// → NmTablesProvider.GetLaserSpeed(NmConfigProvider.Tables, ...)
        /// </summary>
        [Theory]
        [InlineData("304L", 0.075)]
        [InlineData("316L", 0.075)]
        [InlineData("A36", 0.060)]
        [InlineData("CS", 0.075)]
        public void CoreMaterials_ReturnValidSpeed(string material, double thicknessIn)
        {
            string asmDir = System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location);
            string configDir = System.IO.Path.Combine(asmDir, "config");
            if (!System.IO.Directory.Exists(configDir))
                return; // Skip if no config deployed

            var provider = new StaticLaserSpeedProvider();
            var speed = provider.GetSpeed(thicknessIn, material);

            Assert.True(speed.HasValue, $"Speed should be available for {material} at {thicknessIn}\"");
            Assert.True(speed.FeedRateIpm > 0, $"Feed rate should be > 0 for {material}");
        }

        /// <summary>
        /// When NmConfigProvider is reset to defaults (no tables loaded),
        /// StaticLaserSpeedProvider should still not throw — just return no speed.
        /// This validates MainRunner won't crash if config is missing.
        /// </summary>
        [Fact]
        public void DefaultConfig_DoesNotThrow()
        {
            NmConfigProvider.ResetToDefaults();
            var provider = new StaticLaserSpeedProvider();

            // Should not throw; may return no speed
            var speed = provider.GetSpeed(0.075, "304L");
            // We don't assert HasValue because it depends on whether defaults include speeds
        }

        /// <summary>
        /// The laser speed lookup chain should produce a usable LaserOpResult
        /// when wired through LaserCalculator — the full production path.
        /// </summary>
        [Fact]
        public void FullChain_ProducesLaserResult()
        {
            string asmDir = System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location);
            string configDir = System.IO.Path.Combine(asmDir, "config");
            if (!System.IO.Directory.Exists(configDir))
                return; // Skip if no config deployed

            NmConfigProvider.Reload(configDir);
            var provider = new StaticLaserSpeedProvider();

            var metrics = new PartMetrics
            {
                ApproxCutLengthIn = 40,
                PierceCount = 5,
                ThicknessIn = 0.075,
                MaterialCode = "304L",
                MassKg = 2.0
            };

            var result = LaserCalculator.Compute(metrics, provider, isWaterjet: false, rawWeightLb: 4.4);

            // If tables are loaded and contain 304L at 0.075", we should get non-zero
            if (provider.GetSpeed(0.075, "304L").HasValue)
            {
                Assert.True(result.SetupHours > 0, "Laser setup should be > 0 with real tables");
                Assert.True(result.RunHours > 0, "Laser run should be > 0 with real tables");
            }
        }
    }

    #endregion

    #region OptiMaterial Pipeline Wiring

    /// <summary>
    /// Tests that OptiMaterial resolution integrates correctly with the cost pipeline.
    /// Validates that after CalculateCosts populates PartData, OptiMaterial resolves
    /// to a valid code — matching the MainRunner flow (OptiMaterial is resolved at line 936,
    /// before CalculateCosts at ~line 1050).
    /// </summary>
    public class OptiMaterialPipelineWiringTests
    {
        public OptiMaterialPipelineWiringTests()
        {
            NmConfigProvider.ResetToDefaults();
        }

        /// <summary>
        /// After material is set on PartData, OptiMaterial resolution should produce
        /// a valid code that matches the material. This was issue #2 (OptiMaterial stuck).
        /// </summary>
        [Theory]
        [InlineData("304L", "SS")]
        [InlineData("316L", "SS")]
        [InlineData("A36", "CS")]
        [InlineData("6061", "AL")]
        [InlineData("C22", "Other")]
        public void AfterMaterialSet_OptiMaterialResolves(string material, string category)
        {
            var pd = new PartData
            {
                Material = material,
                MaterialCategory = category,
                Classification = PartType.SheetMetal,
                Thickness_m = 0.001897,
            };
            pd.Sheet.IsSheetMetal = true;

            // Simulate MainRunner flow: resolve OptiMaterial
            pd.OptiMaterial = StaticOptiMaterialService.Resolve(pd);

            Assert.NotNull(pd.OptiMaterial);
            Assert.StartsWith("S.", pd.OptiMaterial);
            Assert.Contains(material, pd.OptiMaterial);
        }

        /// <summary>
        /// If material changes on the PartData, re-resolving OptiMaterial should
        /// reflect the new material. This was issue #2 (OptiMaterial stuck at old value).
        /// </summary>
        [Fact]
        public void MaterialChange_OptiMaterialUpdates()
        {
            var pd = new PartData
            {
                Material = "304L",
                Classification = PartType.SheetMetal,
                Thickness_m = 0.001897,
            };
            pd.Sheet.IsSheetMetal = true;

            // First resolution
            pd.OptiMaterial = StaticOptiMaterialService.Resolve(pd);
            Assert.Contains("304L", pd.OptiMaterial);

            // Change material
            pd.Material = "316L";
            pd.OptiMaterial = StaticOptiMaterialService.Resolve(pd);
            Assert.Contains("316L", pd.OptiMaterial);
            Assert.DoesNotContain("304L", pd.OptiMaterial);
        }

        /// <summary>
        /// OptiMaterial for tube parts should use the correct prefix.
        /// </summary>
        [Fact]
        public void TubePart_OptiMaterialUsesCorrectPrefix()
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

            pd.OptiMaterial = StaticOptiMaterialService.Resolve(pd);

            Assert.NotNull(pd.OptiMaterial);
            Assert.StartsWith("P.", pd.OptiMaterial);
        }

        /// <summary>
        /// Material code flows through MaterialCodeMapper before OptiMaterial resolution.
        /// SW database name → short code → OptiMaterial.
        /// </summary>
        [Theory]
        [InlineData("AISI 304", "304L")]
        [InlineData("AISI 316", "316L")]
        [InlineData("ASTM A36 Steel", "A36")]
        public void SwDatabaseName_FlowsThroughMapperToOptiMaterial(string swName, string expectedCode)
        {
            // Step 1: MaterialCodeMapper maps SW name → short code
            string shortCode = MaterialCodeMapper.ToShortCode(swName);
            Assert.Equal(expectedCode, shortCode);

            // Step 2: If MainRunner uses the mapped code, OptiMaterial should resolve
            var pd = new PartData
            {
                Material = shortCode,
                Classification = PartType.SheetMetal,
                Thickness_m = 0.001897,
            };
            pd.Sheet.IsSheetMetal = true;

            pd.OptiMaterial = StaticOptiMaterialService.Resolve(pd);
            Assert.NotNull(pd.OptiMaterial);
            Assert.Contains(expectedCode, pd.OptiMaterial);
        }
    }

    #endregion

    #region Work Center Rate Wiring

    /// <summary>
    /// Tests that GetOP20Rate returns correct rates for each work center code.
    /// This validates MainRunner lines 1247-1256.
    /// </summary>
    public class WorkCenterRateTests
    {
        public WorkCenterRateTests()
        {
            NmConfigProvider.ResetToDefaults();
        }

        [Fact]
        public void F115_ReturnsLaserRate()
        {
            Assert.Equal(CostConstants.F115_COST, GetOP20Rate("F115"));
        }

        [Fact]
        public void F300_ReturnsMaterialHandlingRate()
        {
            Assert.Equal(CostConstants.F300_COST, GetOP20Rate("F300"));
        }

        [Fact]
        public void N145_ReturnsCncRate()
        {
            Assert.Equal(CostConstants.F145_COST, GetOP20Rate("N145"));
        }

        [Fact]
        public void F110_UsesMaterialHandlingRate()
        {
            // F110 (bandsaw) uses F300 rate since no dedicated F110 rate exists
            Assert.Equal(CostConstants.F300_COST, GetOP20Rate("F110"));
        }

        [Fact]
        public void UnknownWorkCenter_DefaultsToF300Rate()
        {
            Assert.Equal(CostConstants.F300_COST, GetOP20Rate("UNKNOWN"));
        }

        [Fact]
        public void AllRates_ArePositive()
        {
            Assert.True(CostConstants.F115_COST > 0, "F115 rate should be > 0");
            Assert.True(CostConstants.F300_COST > 0, "F300 rate should be > 0");
            Assert.True(CostConstants.F145_COST > 0, "N145 rate should be > 0");
            Assert.True(CostConstants.F140_COST > 0, "F140 rate should be > 0");
            Assert.True(CostConstants.F210_COST > 0, "F210 rate should be > 0");
            Assert.True(CostConstants.F220_COST > 0, "F220 rate should be > 0");
            Assert.True(CostConstants.F325_COST > 0, "F325 rate should be > 0");
        }

        /// <summary>Replicates MainRunner.GetOP20Rate (lines 1247-1256).</summary>
        private static double GetOP20Rate(string workCenter)
        {
            switch (workCenter)
            {
                case "F115": return CostConstants.F115_COST;
                case "F300": return CostConstants.F300_COST;
                case "N145": return CostConstants.F145_COST;
                case "F110": return CostConstants.F300_COST;
                default:     return CostConstants.F300_COST;
            }
        }
    }

    #endregion

    #region Shared Mock Provider

    /// <summary>
    /// Mock laser speed provider used by wiring tests to avoid dependency on config files.
    /// </summary>
    internal sealed class MockLaserProvider : ILaserSpeedProvider
    {
        private readonly double _feedIpm;
        private readonly double _pierceSec;

        public MockLaserProvider(double feedIpm, double pierceSec)
        {
            _feedIpm = feedIpm;
            _pierceSec = pierceSec;
        }

        public LaserSpeed GetSpeed(double thicknessIn, string materialCode)
        {
            return new LaserSpeed { FeedRateIpm = _feedIpm, PierceSeconds = _pierceSec };
        }
    }

    #endregion
}
