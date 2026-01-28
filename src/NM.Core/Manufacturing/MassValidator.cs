using System;

namespace NM.Core.Manufacturing
{
    /// <summary>
    /// Validates mass calculations by comparing calculated vs measured values.
    /// Ported from VBA SP.bas CompareMass() function.
    /// </summary>
    public static class MassValidator
    {
        /// <summary>
        /// Default tolerance percentage for mass comparison (5%).
        /// </summary>
        public const double DefaultTolerancePercent = 5.0;

        public sealed class MassComparisonResult
        {
            public double CalculatedMassLb { get; set; }
            public double MeasuredMassLb { get; set; }
            public double PercentDifference { get; set; }
            public bool IsWithinTolerance { get; set; }
            public string Message { get; set; }
        }

        /// <summary>
        /// Compares calculated mass against measured mass.
        /// </summary>
        /// <param name="calculatedMassLb">Mass calculated from geometry (lb).</param>
        /// <param name="measuredMassLb">Mass from SolidWorks mass properties or scale (lb).</param>
        /// <param name="tolerancePercent">Allowable difference percentage (default 5%).</param>
        /// <returns>Comparison result with pass/fail status.</returns>
        public static MassComparisonResult Compare(double calculatedMassLb, double measuredMassLb, double tolerancePercent = DefaultTolerancePercent)
        {
            var result = new MassComparisonResult
            {
                CalculatedMassLb = calculatedMassLb,
                MeasuredMassLb = measuredMassLb
            };

            // Handle edge cases
            if (measuredMassLb <= 0)
            {
                result.IsWithinTolerance = false;
                result.PercentDifference = 100;
                result.Message = "Measured mass is zero or negative - cannot validate";
                return result;
            }

            if (calculatedMassLb <= 0)
            {
                result.IsWithinTolerance = false;
                result.PercentDifference = 100;
                result.Message = "Calculated mass is zero or negative - cannot validate";
                return result;
            }

            // Calculate percent difference from measured value
            result.PercentDifference = Math.Abs((calculatedMassLb - measuredMassLb) / measuredMassLb) * 100;
            result.IsWithinTolerance = result.PercentDifference <= tolerancePercent;

            if (result.IsWithinTolerance)
            {
                result.Message = $"Mass validated: {result.PercentDifference:F1}% difference (within {tolerancePercent}% tolerance)";
            }
            else
            {
                result.Message = $"Mass MISMATCH: {result.PercentDifference:F1}% difference (exceeds {tolerancePercent}% tolerance)";
            }

            return result;
        }

        /// <summary>
        /// Calculates mass from volume and density.
        /// </summary>
        /// <param name="volumeIn3">Volume in cubic inches.</param>
        /// <param name="densityLbPerIn3">Density in lb/in³.</param>
        /// <returns>Mass in pounds.</returns>
        public static double CalculateMassFromVolume(double volumeIn3, double densityLbPerIn3)
        {
            return volumeIn3 * densityLbPerIn3;
        }

        /// <summary>
        /// Calculates mass from blank dimensions and thickness (sheet metal).
        /// </summary>
        /// <param name="lengthIn">Blank length in inches.</param>
        /// <param name="widthIn">Blank width in inches.</param>
        /// <param name="thicknessIn">Material thickness in inches.</param>
        /// <param name="densityLbPerIn3">Density in lb/in³.</param>
        /// <returns>Mass in pounds.</returns>
        public static double CalculateMassFromBlank(double lengthIn, double widthIn, double thicknessIn, double densityLbPerIn3)
        {
            double volumeIn3 = lengthIn * widthIn * thicknessIn;
            return CalculateMassFromVolume(volumeIn3, densityLbPerIn3);
        }

        /// <summary>
        /// Converts kilograms to pounds.
        /// </summary>
        public static double KgToLb(double kg) => kg * 2.20462;

        /// <summary>
        /// Converts pounds to kilograms.
        /// </summary>
        public static double LbToKg(double lb) => lb / 2.20462;
    }
}
