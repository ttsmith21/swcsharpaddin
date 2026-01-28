using System;

namespace NM.Core.Manufacturing
{
    /// <summary>
    /// Calculates material costs for sheet metal and tube parts.
    /// Ported from VBA modMaterialCost.bas MaterialCost() function.
    /// </summary>
    public static class MaterialCostCalculator
    {
        /// <summary>
        /// Material pricing per pound by category.
        /// </summary>
        public static class MaterialPricing
        {
            public const double Stainless304_PerLb = 1.75;
            public const double Stainless316_PerLb = 2.25;
            public const double CarbonSteel_PerLb = 0.55;
            public const double Aluminum6061_PerLb = 2.50;
            public const double Aluminum5052_PerLb = 2.35;
            public const double GalvanizedPerLb = 0.65;

            /// <summary>
            /// Gets material cost per pound based on material code.
            /// </summary>
            public static double GetCostPerLb(string materialCode)
            {
                if (string.IsNullOrWhiteSpace(materialCode))
                    return CarbonSteel_PerLb;

                var m = materialCode.ToUpperInvariant();

                if (m.Contains("316")) return Stainless316_PerLb;
                if (m.Contains("304") || m.Contains("309") || m.Contains("SS")) return Stainless304_PerLb;
                // Check GALV before AL to avoid "GALV" matching "AL"
                if (m.Contains("GALV") || m.Contains("GV")) return GalvanizedPerLb;
                if (m.Contains("6061")) return Aluminum6061_PerLb;
                if (m.Contains("5052") || m.Contains("AL")) return Aluminum5052_PerLb;
                if (m.Contains("A36") || m.Contains("CS") || m.Contains("CARBON")) return CarbonSteel_PerLb;

                return CarbonSteel_PerLb;
            }
        }

        public sealed class MaterialCostInput
        {
            public double WeightLb { get; set; }
            public string MaterialCode { get; set; }
            public int Quantity { get; set; } = 1;
            public double? OverrideCostPerLb { get; set; }
            public double NestEfficiency { get; set; } = 1.0; // 1.0 = no waste, 0.75 = 75% efficient
        }

        public sealed class MaterialCostResult
        {
            public double CostPerLb { get; set; }
            public double WeightLb { get; set; }
            public double RawWeightLb { get; set; } // Before nest efficiency adjustment
            public double CostPerPiece { get; set; }
            public double TotalMaterialCost { get; set; }
            public string MaterialCode { get; set; }
        }

        /// <summary>
        /// Calculates material cost for a part.
        /// </summary>
        public static MaterialCostResult Calculate(MaterialCostInput input)
        {
            if (input == null)
                return new MaterialCostResult();

            double costPerLb = input.OverrideCostPerLb ?? MaterialPricing.GetCostPerLb(input.MaterialCode);
            double rawWeight = input.WeightLb;

            // Apply nest efficiency - if efficiency is 75%, we're using 133% of material
            double adjustedWeight = input.NestEfficiency > 0 ? rawWeight / input.NestEfficiency : rawWeight;

            double costPerPiece = adjustedWeight * costPerLb;
            int qty = Math.Max(1, input.Quantity);

            return new MaterialCostResult
            {
                CostPerLb = costPerLb,
                WeightLb = adjustedWeight,
                RawWeightLb = rawWeight,
                CostPerPiece = costPerPiece,
                TotalMaterialCost = costPerPiece * qty,
                MaterialCode = input.MaterialCode
            };
        }

        /// <summary>
        /// Calculates raw material weight for a sheet metal blank.
        /// </summary>
        /// <param name="lengthIn">Blank length in inches.</param>
        /// <param name="widthIn">Blank width in inches.</param>
        /// <param name="thicknessIn">Material thickness in inches.</param>
        /// <param name="materialCode">Material code for density lookup.</param>
        /// <returns>Weight in pounds.</returns>
        public static double CalculateBlankWeight(double lengthIn, double widthIn, double thicknessIn, string materialCode)
        {
            double volumeIn3 = lengthIn * widthIn * thicknessIn;
            double density = Rates.GetDensityLbPerIn3(materialCode);
            return volumeIn3 * density;
        }

        /// <summary>
        /// Calculates raw material weight for a tube/pipe.
        /// </summary>
        /// <param name="odIn">Outside diameter in inches.</param>
        /// <param name="idIn">Inside diameter in inches (OD - 2*wall).</param>
        /// <param name="lengthIn">Tube length in inches.</param>
        /// <param name="materialCode">Material code for density lookup.</param>
        /// <returns>Weight in pounds.</returns>
        public static double CalculateTubeWeight(double odIn, double idIn, double lengthIn, string materialCode)
        {
            // Volume = π * L * (OD² - ID²) / 4
            double volumeIn3 = Math.PI * lengthIn * (odIn * odIn - idIn * idIn) / 4.0;
            double density = Rates.GetDensityLbPerIn3(materialCode);
            return volumeIn3 * density;
        }

        /// <summary>
        /// Calculates raw material weight for round bar stock.
        /// </summary>
        /// <param name="diameterIn">Bar diameter in inches.</param>
        /// <param name="lengthIn">Bar length in inches.</param>
        /// <param name="materialCode">Material code for density lookup.</param>
        /// <returns>Weight in pounds.</returns>
        public static double CalculateRoundBarWeight(double diameterIn, double lengthIn, string materialCode)
        {
            // Volume = π * D² * L / 4
            double volumeIn3 = Math.PI * diameterIn * diameterIn * lengthIn / 4.0;
            double density = Rates.GetDensityLbPerIn3(materialCode);
            return volumeIn3 * density;
        }
    }
}
