using System;

namespace NM.Core.Manufacturing
{
    public sealed class F325Calculator
    {
        // Use CostConstants.F325_COST for hourly rate
        private const double ROLL_SETUP_HOURS = 0.5; // hours
        private const double ROLL_SPEED_FPM = 10.0;  // feet per minute

        public sealed class F325Result
        {
            public bool RequiresRollForming { get; set; }
            public double SetupHours { get; set; }
            public double RunHours { get; set; }
            public double TotalCost { get; set; }
            public string Notes { get; set; }
        }

        public F325Result CalculateRollForming(double maxRadiusInches, double arcLengthInches, int quantity)
        {
            var r = new F325Result();
            if (maxRadiusInches <= 2.0)
            {
                r.RequiresRollForming = false; r.SetupHours = 0; r.RunHours = 0; r.TotalCost = 0; return r;
            }

            r.RequiresRollForming = true;
            r.Notes = $"Roll forming required - radius {maxRadiusInches:0.###}\" > 2\"";

            r.SetupHours = ROLL_SETUP_HOURS;
            double arcFeet = Math.Max(0.0, arcLengthInches) / 12.0;
            double runMinutes = (ROLL_SPEED_FPM > 0) ? (arcFeet / ROLL_SPEED_FPM) * 60.0 : 0.0;
            r.RunHours = runMinutes / 60.0;

            double totalHours = r.SetupHours + (r.RunHours * Math.Max(1, quantity));
            r.TotalCost = totalHours * CostConstants.F325_COST;
            return r;
        }

        public static double CalculateArcLength(double radiusInches, double angleRadians)
            => Math.Abs(radiusInches * angleRadians);
    }
}
