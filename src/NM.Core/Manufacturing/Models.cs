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
}
