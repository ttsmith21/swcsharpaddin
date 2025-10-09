using System;

namespace NM.Core.Manufacturing
{
    public sealed class F220Input
    {
        public int Setups { get; set; }
        public int Holes { get; set; }
    }

    public sealed class F220Result
    {
        public double SetupHours { get; set; }
        public double RunHours { get; set; }
        // Price is computed by caller using CostConstants.F220_COST (not present in some versions)
    }

    public static class F220Calculator
    {
        public static F220Result Compute(F220Input input)
        {
            var r = new F220Result();
            if (input == null) input = new F220Input();

            // Setup hours = setups × 0.015 + 0.085, min 0.1 hours
            double setup = (input.Setups * 0.015) + 0.085;
            if (setup < 0.1) setup = 0.1;
            r.SetupHours = setup;

            // Run hours = holes × 0.01
            r.RunHours = Math.Max(0, input.Holes) * 0.01;
            return r;
        }
    }
}
