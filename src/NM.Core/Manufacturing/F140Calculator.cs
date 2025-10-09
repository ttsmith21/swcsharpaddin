using System;

namespace NM.Core.Manufacturing
{
    public sealed class F140Result
    {
        public double SetupHours { get; set; }
        public double RunHours { get; set; }
        public double TotalHours(int quantity) => SetupHours + Math.Max(1, quantity) * RunHours;
        public double Price(int quantity) => TotalHours(quantity) * CostConstants.F140_COST;
    }

    public static class F140Calculator
    {
        // Exact legacy logic
        public static F140Result Compute(BendInfo bend, double weightLb, int quantity)
        {
            var r = new F140Result();
            if (bend == null) bend = new BendInfo();

            // Setup minutes = LongestBend(ft) * 1.25 + 10
            double setupMinutes = (bend.LongestBendIn / 12.0) * 1.25 + 10.0;
            r.SetupHours = Math.Max(0.0, setupMinutes / 60.0);

            // Bend rate selection (seconds per bend)
            int rateSec;
            if (weightLb > 100.0) rateSec = 400;
            else if (weightLb > 40.0) rateSec = 200;
            else if (weightLb > 5.0 || bend.LongestBendIn > 60.0) rateSec = 45;
            else if (bend.LongestBendIn > 12.0) rateSec = 30;
            else rateSec = 10;

            int bendOps = Math.Max(0, bend.Count) + (bend.NeedsFlip ? 1 : 0);
            double runSeconds = bendOps * rateSec;
            r.RunHours = Math.Max(0.0, runSeconds / 3600.0);

            return r;
        }
    }
}
