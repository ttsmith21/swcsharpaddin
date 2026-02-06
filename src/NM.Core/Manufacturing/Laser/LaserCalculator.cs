using System;

namespace NM.Core.Manufacturing.Laser
{
    public sealed class LaserOpResult
    {
        public double SetupHours { get; set; }
        public double RunHours { get; set; }
        public double TotalHours => SetupHours + RunHours;
        public double Cost => TotalHours * CostConstants.F115_COST;
    }

    public static class LaserCalculator
    {
        // Legacy constants
        private const double LASER_SETUP_RATE_MIN = 5.0;   // minutes per sheet
        private const double LASER_SETUP_TIME_MIN = 0.5;   // fixed minutes (laser)
        private const double WATERJET_SETUP_TIME_MIN = 15.0; // fixed minutes (waterjet)
        private const double WATERJET_SETUP_RATE_MIN = 30.0; // minutes per sheet (waterjet)
        private const double MIN_SETUP_HOURS = 0.01;       // minimum setup hours

        // Backward-compat shim: uses mass for rawWeight and assumes laser (not waterjet)
        public static LaserOpResult Compute(PartMetrics m, ILaserSpeedProvider provider)
        {
            double rawWeightLb = (m != null && m.MassKg > 0) ? m.MassKg * 2.20462262185 : 0.0;
            return Compute(m, provider, isWaterjet: false, rawWeightLb: rawWeightLb);
        }

        // Exact behavior per legacy VBA
        public static LaserOpResult Compute(PartMetrics m, ILaserSpeedProvider provider, bool isWaterjet, double rawWeightLb)
        {
            var res = new LaserOpResult();
            if (m == null || provider == null) return res;

            var speed = provider.GetSpeed(m.ThicknessIn, m.MaterialCode);
            if (!speed.HasValue) return res;

            // total pierces are provided (loops + 2 + floor(length/30))
            double totalPierces = Math.Max(0, m.PierceCount);
            double totalPierceSeconds = isWaterjet ? 0.0 : (totalPierces * Math.Max(0, speed.PierceSeconds));

            // Cut time in minutes = length / IPM
            double cutMinutes = (Math.Max(0, m.ApproxCutLengthIn) / Math.Max(1e-9, speed.FeedRateIpm));

            // Sheet weight (for setup proportion)
            double rho = Rates.GetDensityLbPerIn3(m.MaterialCode);
            double sheetWeightLb = m.ThicknessIn * 60.0 * 120.0 * rho;

            // Setup time: fixed only (VBA adds proportional term to run, not setup)
            double setupMinutes = isWaterjet ? WATERJET_SETUP_TIME_MIN : LASER_SETUP_TIME_MIN;
            double setupHours = Math.Max(MIN_SETUP_HOURS, setupMinutes / 60.0);

            // Run time: pierce + cut + proportional sheet-loading time
            double proportionalMinutes = 0.0;
            if (sheetWeightLb > 0 && rawWeightLb > 0)
            {
                double rate = isWaterjet ? WATERJET_SETUP_RATE_MIN : LASER_SETUP_RATE_MIN;
                proportionalMinutes = (rawWeightLb / sheetWeightLb) * rate;
            }
            double runHours = (totalPierceSeconds / 3600.0) + ((cutMinutes + proportionalMinutes) / 60.0);

            res.SetupHours = setupHours;
            res.RunHours = runHours;
            return res;
        }
    }
}
