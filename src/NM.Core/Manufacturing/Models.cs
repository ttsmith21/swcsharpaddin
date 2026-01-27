using System;
using NM.Core;

namespace NM.Core.Manufacturing
{
    // Input metrics extracted from SolidWorks (units in comments)
    public sealed class PartMetrics
    {
        public double MassKg { get; set; }            // kg (0 if unavailable)
        public double VolumeM3 { get; set; }          // m^3
        public double ThicknessIn { get; set; }       // inches
        public string MaterialCode { get; set; }      // e.g., 304L, 316L, A36, 6061
        public int BendCount { get; set; }            // bends (sheet metal)
        public double ApproxCutLengthIn { get; set; } // inches
        public int PierceCount { get; set; }          // count

        // Legacy properties used by VBA logic
        public string WeightCalcMode { get; set; }    // rbWeightCalc: "0"=efficiency, "1"=manual
        public double NestEfficiencyPercent { get; set; } // e.g., 75.0 means 75%
        public double BlankLengthIn { get; set; }     // inches (manual mode)
        public double BlankWidthIn { get; set; }      // inches (manual mode)

        // Job context
        public int Quantity { get; set; }             // pieces
        public DifficultyLevel Difficulty { get; set; } // tolerance multipliers
    }

    // Options controlling calculators
    public sealed class CalcOptions
    {
        public bool UseMassIfAvailable { get; set; } = true;
        public bool QuoteEnabled { get; set; } = false;
    }

    // Output of the calculator
    public sealed class CalcResult
    {
        // Weight
        public double WeightLb { get; set; }      // pounds
        public double RawWeightLb { get; set; }   // raw weight before sheet percent (legacy RawWeight)
        public double SheetPercent { get; set; }  // fraction (0..1)

        // Time summaries (rough v1 or per-legacy ops)
        public double LaserMinutes { get; set; }  // F115 minutes (rough)
        public double BrakeMinutes { get; set; }  // F140 minutes (rough)

        public string Notes { get; set; }
    }

    // Total cost calculation inputs
    public sealed class TotalCostInputs
    {
        public int Quantity { get; set; }
        public double RawWeightLb { get; set; }
        public double MaterialCostPerLb { get; set; }
        public double F115Price { get; set; }  // Laser
        public double F140Price { get; set; }  // Brake
        public double F220Price { get; set; }
        public double F325Price { get; set; }
        public DifficultyLevel Difficulty { get; set; }
    }

    // Total cost calculation result
    public sealed class TotalCostResult
    {
        public int Quantity { get; set; }
        public double MaterialCostPerLB { get; set; }
        public double TotalMaterialCost { get; set; }
        public double TotalProcessingCost { get; set; }
        public double TotalCost { get; set; }
    }

    // Simple total cost calculator
    public static class TotalCostCalculator
    {
        public static TotalCostResult Compute(TotalCostInputs inputs)
        {
            if (inputs == null) return new TotalCostResult();

            // Material cost = weight * cost per lb * qty
            double matCost = inputs.RawWeightLb * inputs.MaterialCostPerLb * inputs.Quantity;

            // Processing cost = sum of operations * difficulty multiplier
            double diffMult = inputs.Difficulty == DifficultyLevel.Tight ? 1.2
                            : inputs.Difficulty == DifficultyLevel.Loose ? 0.9
                            : 1.0;
            double procCost = (inputs.F115Price + inputs.F140Price + inputs.F220Price + inputs.F325Price) * inputs.Quantity * diffMult;

            return new TotalCostResult
            {
                Quantity = inputs.Quantity,
                MaterialCostPerLB = inputs.MaterialCostPerLb,
                TotalMaterialCost = matCost,
                TotalProcessingCost = procCost,
                TotalCost = matCost + procCost
            };
        }
    }
}
