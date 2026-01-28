using System;

namespace NM.Core.Manufacturing
{
    /// <summary>
    /// Bend tonnage calculator - ported from VBA modMaterialCost.bas CheckBendTonnage().
    /// Validates if a bend can be performed on available press brakes.
    /// </summary>
    public static class BendTonnageCalculator
    {
        // Press brake capacity constants (tons)
        public const double SMALL_BRAKE_TONS = 90;
        public const double MEDIUM_BRAKE_TONS = 175;
        public const double LARGE_BRAKE_TONS = 350;

        /// <summary>
        /// Calculates required tonnage for a bend.
        /// Formula: Tonnage = (575 × Thickness² × BendLength) / DieOpening
        /// </summary>
        /// <param name="thicknessIn">Material thickness in inches</param>
        /// <param name="bendLengthIn">Bend length in inches</param>
        /// <param name="tensileStrength">Material tensile strength (psi). Default 60000 for 304 SS</param>
        /// <returns>Required tonnage</returns>
        public static double CalculateTonnage(double thicknessIn, double bendLengthIn, double tensileStrength = 60000)
        {
            if (thicknessIn <= 0 || bendLengthIn <= 0)
                return 0;

            // Die opening rule of thumb: 8× thickness
            double dieOpening = thicknessIn * 8;
            if (dieOpening < 0.25)
                dieOpening = 0.25; // Minimum practical die opening

            // Tonnage formula: (575 × T² × L × TensileRatio) / V
            // Where TensileRatio = TensileStrength / 60000 (normalized to 304 SS)
            double tensileRatio = tensileStrength / 60000.0;
            double tonnage = (575 * thicknessIn * thicknessIn * bendLengthIn * tensileRatio) / dieOpening;

            return tonnage;
        }

        /// <summary>
        /// Checks if a bend can be performed and returns the required brake.
        /// </summary>
        public static BendTonnageResult CheckBend(double thicknessIn, double bendLengthIn, string material)
        {
            var result = new BendTonnageResult();
            result.ThicknessIn = thicknessIn;
            result.BendLengthIn = bendLengthIn;

            // Get tensile strength based on material
            double tensile = GetTensileStrength(material);
            result.TensileStrength = tensile;

            // Calculate required tonnage
            result.RequiredTonnage = CalculateTonnage(thicknessIn, bendLengthIn, tensile);

            // Determine which brake can handle it
            if (result.RequiredTonnage <= SMALL_BRAKE_TONS)
            {
                result.CanBend = true;
                result.RecommendedBrake = "Small (90T)";
                result.WorkCenter = "F140";
            }
            else if (result.RequiredTonnage <= MEDIUM_BRAKE_TONS)
            {
                result.CanBend = true;
                result.RecommendedBrake = "Medium (175T)";
                result.WorkCenter = "F140";
            }
            else if (result.RequiredTonnage <= LARGE_BRAKE_TONS)
            {
                result.CanBend = true;
                result.RecommendedBrake = "Large (350T)";
                result.WorkCenter = "F145"; // CNC bending for heavy work
            }
            else
            {
                result.CanBend = false;
                result.RecommendedBrake = "OUTSOURCE";
                result.WorkCenter = "OS";
                result.Message = $"Bend requires {result.RequiredTonnage:F0}T, exceeds max capacity of {LARGE_BRAKE_TONS}T";
            }

            return result;
        }

        /// <summary>
        /// Gets approximate tensile strength for material.
        /// </summary>
        public static double GetTensileStrength(string material)
        {
            if (string.IsNullOrEmpty(material))
                return 60000; // Default to 304 SS

            var m = material.ToUpperInvariant();

            // Stainless steel grades
            if (m.Contains("304")) return 75000;
            if (m.Contains("316")) return 75000;
            if (m.Contains("309")) return 75000;
            if (m.Contains("2205")) return 95000;  // Duplex is stronger
            if (m.Contains("AL6XN")) return 100000;
            if (m.Contains("C22") || m.Contains("HASTELLOY")) return 100000;

            // Carbon steel
            if (m.Contains("A36") || m.Contains("CS") || m.Contains("CARBON"))
                return 58000;

            // Aluminum
            if (m.Contains("6061")) return 45000;
            if (m.Contains("5052")) return 33000;
            if (m.Contains("AL")) return 40000;

            return 60000; // Default
        }
    }

    public sealed class BendTonnageResult
    {
        public double ThicknessIn { get; set; }
        public double BendLengthIn { get; set; }
        public double TensileStrength { get; set; }
        public double RequiredTonnage { get; set; }
        public bool CanBend { get; set; }
        public string RecommendedBrake { get; set; }
        public string WorkCenter { get; set; }
        public string Message { get; set; }
    }
}
