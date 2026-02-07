namespace NM.Core.Config.Sections
{
    /// <summary>
    /// Manufacturing parameters: press brake rates, laser setup, waterjet setup, sheet dimensions.
    /// Loaded from nm-config.json.
    /// </summary>
    public sealed class ManufacturingParams
    {
        public PressBrakeParams PressBrake { get; set; } = new PressBrakeParams();
        public LaserParams Laser { get; set; } = new LaserParams();
        public WaterjetParams Waterjet { get; set; } = new WaterjetParams();
        public StandardSheetParams StandardSheet { get; set; } = new StandardSheetParams();
        public PressBrakeCapacityParams PressBrakeCapacity { get; set; } = new PressBrakeCapacityParams();
        public DeburringParams Deburring { get; set; } = new DeburringParams();
        public RollFormingParams RollForming { get; set; } = new RollFormingParams();
        public TappingParams Tapping { get; set; } = new TappingParams();
        public CuttingParams Cutting { get; set; } = new CuttingParams();
    }

    public sealed class PressBrakeParams
    {
        /// <summary>Seconds per bend for each rate tier [Rate1..Rate5].</summary>
        public double[] RateSeconds { get; set; } = { 10, 30, 45, 200, 400 };
        /// <summary>Weight thresholds in lbs for rate tier selection [Rate1Max, Rate2Max, Rate3Max].</summary>
        public double[] WeightThresholdsLbs { get; set; } = { 5, 40, 100 };
        /// <summary>Length thresholds in inches for rate tier selection [Rate1Max, Rate2Max].</summary>
        public double[] LengthThresholdsIn { get; set; } = { 12, 60 };
        public double SetupMinutesPerFoot { get; set; } = 1.25;
        public double SetupFixedMinutes { get; set; } = 10.0;
    }

    public sealed class PressBrakeCapacityParams
    {
        public double SmallBrakeTons { get; set; } = 90.0;
        public double MediumBrakeTons { get; set; } = 175.0;
        public double LargeBrakeTons { get; set; } = 350.0;
        public double DieOpeningMultiplier { get; set; } = 8.0;
        public double MinDieOpeningIn { get; set; } = 0.25;
        public double TonnageFormulaNumerator { get; set; } = 575.0;
        public double DefaultTensileStrength { get; set; } = 60000.0;
    }

    public sealed class LaserParams
    {
        public double SetupMinutesPerSheet { get; set; } = 5.0;
        public double SetupFixedMinutes { get; set; } = 0.5;
        public double MinSetupHours { get; set; } = 0.01;
        public double TabSpacingIn { get; set; } = 30.0;
    }

    public sealed class WaterjetParams
    {
        public double SetupFixedMinutes { get; set; } = 15.0;
        public double SetupMinutesPerLoad { get; set; } = 30.0;
    }

    public sealed class StandardSheetParams
    {
        public double WidthIn { get; set; } = 60.0;
        public double LengthIn { get; set; } = 120.0;
    }

    public sealed class DeburringParams
    {
        public double RateInchesPerMinute { get; set; } = 60.0;
    }

    public sealed class RollFormingParams
    {
        public double SetupHours { get; set; } = 0.5;
        public double SpeedFeetPerMinute { get; set; } = 10.0;
        public double MinRadiusForRolling { get; set; } = 2.0;
    }

    public sealed class TappingParams
    {
        public double SetupPerHoleHours { get; set; } = 0.015;
        public double SetupFixedHours { get; set; } = 0.085;
        public double MinSetupHours { get; set; } = 0.1;
        public double RunHoursPerHole { get; set; } = 0.01;
    }

    public sealed class CuttingParams
    {
        public double PierceConstant { get; set; } = 2.0;
        public int TabSpacing { get; set; } = 30;
    }
}
