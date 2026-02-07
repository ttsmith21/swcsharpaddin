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
        // VBA constants from modMaterialCost.bas
        private const double Rate1 = 10;   // seconds - light parts, short bends
        private const double Rate2 = 30;   // seconds
        private const double Rate3 = 45;   // seconds
        private const double Rate4 = 200;  // seconds
        private const double Rate5 = 400;  // seconds - very heavy parts

        private const double Rate3Weight = 100; // lbs - max weight for rate 3
        private const double Rate2Weight = 40;  // lbs - max weight for rate 2
        private const double Rate1Weight = 5;   // lbs - max weight for rate 1
        private const double Rate1Length = 12;  // inches - max length for rate 1
        private const double Rate2Length = 60;  // inches - max length for rate 2

        private const double SetupRate = 1.25;  // minutes per foot for brake setup
        private const double BrakeSetup = 10;   // minutes - brake setup constant

        /// <summary>
        /// Compute F140 press brake setup and run time.
        /// Matches VBA CalculateBendInfo exactly.
        /// </summary>
        /// <param name="bend">Bend info (count, longest bend line, needs flip).</param>
        /// <param name="weightLb">Part weight in lbs (VBA: GetMass).</param>
        /// <param name="partLengthIn">Longest flat face dimension in inches (VBA: dblLength1 from LengthWidth).
        /// This is the part's overall length, NOT the longest bend line. Used for rate lookup.</param>
        /// <param name="quantity">Quantity for total cost calculation.</param>
        public static F140Result Compute(BendInfo bend, double weightLb, double partLengthIn, int quantity)
        {
            var r = new F140Result();
            if (bend == null) bend = new BendInfo();

            // VBA: dblF140_S = dblLongestBend * cdblSetupRate / 12 + cdblBreakSetup
            // Setup uses the longest BEND LINE (not part length)
            double setupMinutes = bend.LongestBendIn * SetupRate / 12.0 + BrakeSetup;
            r.SetupHours = Math.Max(0.0, setupMinutes / 60.0);

            // VBA: dblRunRate = FindRate(dblWeight, dblLength1)
            // Rate lookup uses part WEIGHT and part LENGTH (longest flat face dimension)
            double rateSec = FindRate(weightLb, partLengthIn);

            // VBA: If blnFlip Then dblF140_R = dblRunRate * (intBendCount + 1)
            //      Else dblF140_R = dblRunRate * intBendCount
            int bendOps = Math.Max(0, bend.Count) + (bend.NeedsFlip ? 1 : 0);
            double runSeconds = bendOps * rateSec;
            r.RunHours = Math.Max(0.0, runSeconds / 3600.0);

            return r;
        }

        /// <summary>
        /// VBA FindRate function — exact port.
        /// Determines seconds per bend based on part weight (lbs) and part length (inches).
        /// </summary>
        private static double FindRate(double weightLb, double lengthIn)
        {
            // VBA modMaterialCost.bas lines 556-568
            if (weightLb > Rate3Weight)                              // > 100 lbs
                return Rate5;                                        // 400 sec/bend
            else if (weightLb > Rate2Weight)                         // > 40 lbs
                return Rate4;                                        // 200 sec/bend
            else if (weightLb > Rate1Weight || lengthIn > Rate2Length) // > 5 lbs OR > 60 in
                return Rate3;                                        // 45 sec/bend
            else if (weightLb > Rate1Weight || lengthIn > Rate1Length) // > 5 lbs OR > 12 in
                return Rate2;                                        // 30 sec/bend
            else
                return Rate1;                                        // 10 sec/bend
        }
    }
}
