using System;

namespace NM.Core.Manufacturing
{
    /// <summary>
    /// F210 Deburr calculator - ported from VBA SP.bas N210() function.
    /// Calculates deburr time based on cut perimeter.
    /// </summary>
    public static class F210Calculator
    {
        // DeburRate from VBA: 3600 = 60 inches per minute converted to hours
        // This means we process 60 inches per minute = 1 inch per second
        private const double DEBURR_RATE_IN_PER_MIN = 60.0;

        /// <summary>
        /// Calculates deburr time in hours based on cut perimeter.
        /// </summary>
        /// <param name="cutPerimeterInches">Total cut perimeter in inches</param>
        /// <returns>Deburr time in hours</returns>
        public static double ComputeHours(double cutPerimeterInches)
        {
            if (cutPerimeterInches <= 0)
                return 0;

            // Time in minutes = perimeter / rate
            double minutes = cutPerimeterInches / DEBURR_RATE_IN_PER_MIN;
            return minutes / 60.0;
        }

        /// <summary>
        /// Calculates deburr cost based on cut perimeter.
        /// </summary>
        /// <param name="cutPerimeterInches">Total cut perimeter in inches</param>
        /// <param name="quantity">Number of parts</param>
        /// <returns>Total deburr cost</returns>
        public static double ComputeCost(double cutPerimeterInches, int quantity = 1)
        {
            double hours = ComputeHours(cutPerimeterInches);
            return hours * Math.Max(1, quantity) * CostConstants.F210_COST;
        }
    }

    /// <summary>
    /// F300 Material handling calculator.
    /// Calculates handling time based on part weight and length.
    /// </summary>
    public static class F300Calculator
    {
        /// <summary>
        /// Determines handling rate based on part weight.
        /// </summary>
        public static double GetHandlingRate(double weightLb)
        {
            // Rate tiers from VBA FindRate function
            if (weightLb > 100) return 400;  // Heavy parts need crane
            if (weightLb > 40) return 200;   // Two-person lift
            if (weightLb > 5) return 45;     // Careful handling
            return 10;                        // Light parts
        }

        /// <summary>
        /// Calculates handling time in hours.
        /// </summary>
        public static double ComputeHours(double weightLb, double lengthFt)
        {
            // Adjust rate if part is long
            double baseRate = GetHandlingRate(weightLb);
            if (lengthFt > 5 && weightLb <= 5)
                baseRate = Math.Max(baseRate, 30);

            return baseRate / 3600.0; // Convert seconds to hours
        }

        /// <summary>
        /// Calculates handling cost.
        /// </summary>
        public static double ComputeCost(double weightLb, double lengthFt, int quantity = 1)
        {
            double hours = ComputeHours(weightLb, lengthFt);
            return hours * Math.Max(1, quantity) * CostConstants.F300_COST;
        }
    }
}
