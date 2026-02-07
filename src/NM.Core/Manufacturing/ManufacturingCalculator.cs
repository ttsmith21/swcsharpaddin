using System;
using NM.Core;
using static NM.Core.Constants.UnitConversions;

namespace NM.Core.Manufacturing
{
    public static class ManufacturingCalculator
    {
        // Standard sheet constants (inches)
        private const double STANDARD_SHEET_WIDTH = 60.0;
        private const double STANDARD_SHEET_LENGTH = 120.0;

        // Thickness correction multipliers (legacy exact thresholds)
        private static double GetEfficiencyMultiplier(double tIn)
        {
            if (tIn > 1.2) return 1.015;
            if (tIn > 0.9) return 1.022;
            if (tIn > 0.82) return 1.022;
            if (tIn > 0.7) return 1.026;
            if (tIn > 0.59) return 1.028;
            if (tIn > 0.43) return 1.037;
            if (tIn > 0.34) return 1.054;
            if (tIn > 0.29) return 1.054;
            if (tIn > 0.23) return 1.069;
            if (tIn > 0.1875) return 1.096;
            return 1.0;
        }

        public static CalcResult Compute(PartMetrics m, CalcOptions opt)
        {
            var res = new CalcResult { Notes = string.Empty };
            if (m == null) return res;

            // Determine density (lb/in^3) from material
            double rho = Rates.GetDensityLbPerIn3(m.MaterialCode);
            double t = m.ThicknessIn;
            double mult = GetEfficiencyMultiplier(t);

            // Weight calculation modes per legacy (rbWeightCalc)
            // "0" = efficiency using NestEfficiency; "1" = manual LxW
            bool efficiencyMode = (m.WeightCalcMode ?? "0") == "0";

            double rawWeightLb = 0.0;
            if (efficiencyMode)
            {
                // Efficiency mode: use mass if available else volume; adjust by nest efficiency and thickness multiplier
                if ((opt?.UseMassIfAvailable != false) && m.MassKg > 0)
                {
                    double massLb = m.MassKg * KgToLbs;
                    // VBA: rawWeight = (partMass / nestEfficiency) * 100 * multiplier
                    double eff = (m.NestEfficiencyPercent > 0) ? m.NestEfficiencyPercent : 100.0;
                    rawWeightLb = (massLb / eff) * 100.0 * mult;
                }
                else if (m.VolumeM3 > 0)
                {
                    double volIn3 = m.VolumeM3 * (MetersToInches * MetersToInches * MetersToInches);
                    double eff = (m.NestEfficiencyPercent > 0) ? m.NestEfficiencyPercent : 100.0;
                    // Convert volume*density to lb, then divide by efficiency percent and apply multiplier
                    double partLb = volIn3 * rho;
                    rawWeightLb = (partLb / eff) * 100.0 * mult;
                }
            }
            else
            {
                // Manual mode: thickness * length * width * density * multiplier
                double L = m.BlankLengthIn;
                double W = m.BlankWidthIn;
                rawWeightLb = t * L * W * rho * mult;
            }

            // Sheet percentage
            double sheetLb = t * STANDARD_SHEET_WIDTH * STANDARD_SHEET_LENGTH * rho;
            double sheetPercent = (sheetLb > 0) ? (rawWeightLb / sheetLb) : 0.0;

            // Output
            res.RawWeightLb = RoundTo(rawWeightLb, 4);
            res.SheetPercent = RoundTo(sheetPercent, 4);

            // Preferred WeightLb: if direct mass available, use that; else use rawWeight
            double directLb = (m.MassKg > 0) ? m.MassKg * KgToLbs : 0.0;
            res.WeightLb = RoundTo(directLb > 0 ? directLb : rawWeightLb, 3);

            // Placeholders: F115/F140 minutes will be computed by dedicated modules later (laser, brake)
            res.LaserMinutes = 0;
            res.BrakeMinutes = 0;

            return res;
        }

        private static double RoundTo(double v, int digits)
        {
            return System.Math.Round(v, digits, System.MidpointRounding.AwayFromZero);
        }
    }
}
