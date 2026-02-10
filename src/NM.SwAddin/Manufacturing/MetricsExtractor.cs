using System;
using NM.Core;
using NM.Core.Manufacturing;
using SolidWorks.Interop.sldworks;
using NM.SwAddin; // for SolidWorksApiWrapper

namespace NM.SwAddin.Manufacturing
{
    // Extracts metrics from SolidWorks safely (COM use lives here)
    public static class MetricsExtractor
    {
        private const double DEFAULT_NEST_EFFICIENCY = 80.0; // percent (SOLIDWORKS template default)

        public static PartMetrics FromModel(IModelDoc2 doc, ModelInfo info)
        {
            // Material: read from SolidWorks material assignment.
            // rbMaterialType is "0"/"1"/"2" (Tab Builder radio button), NOT a material code.
            string materialCode;
            try { materialCode = SolidWorksApiWrapper.GetMaterialName(doc); } catch { materialCode = string.Empty; }

            // Legacy custom properties
            string rbWeightCalc = info?.CustomProperties?.GetPropertyValue("rbWeightCalc")?.ToString() ?? "0";
            double nestEff = TryParseDouble(info?.CustomProperties?.GetPropertyValue("NestEfficiency")?.ToString());
            if (nestEff <= 0)
            {
                nestEff = DEFAULT_NEST_EFFICIENCY;
                try
                {
                    info?.CustomProperties?.SetPropertyValue("NestEfficiency", nestEff.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture), CustomPropertyType.Number);
                }
                catch { }
            }
            double blankL = TryParseDouble(info?.CustomProperties?.GetPropertyValue("Length")?.ToString());
            double blankW = TryParseDouble(info?.CustomProperties?.GetPropertyValue("Width")?.ToString());
            int qty = (int)TryParseDouble(info?.CustomProperties?.GetPropertyValue("Quantity")?.ToString());
            var diff = DifficultyLevel.Normal;
            var diffStr = info?.CustomProperties?.GetPropertyValue("rbWeightTolerance")?.ToString();
            if (!string.IsNullOrWhiteSpace(diffStr))
            {
                var u = diffStr.ToUpperInvariant();
                if (u.Contains("TIGHT")) diff = DifficultyLevel.Tight;
                else if (u.Contains("LOOSE")) diff = DifficultyLevel.Loose;
            }

            // Has internal cuts heuristic (pierce > 0 or internal length > 0). Placeholder for future pierce/cut detection.
            // TODO: Implement internal cut detection based on pierce count or internal cut length
            _ = false; // hasInternalCuts - placeholder, suppress warning

            // WARN if thickness is missing - this will corrupt downstream calculations
            double thicknessIn = info?.ThicknessInInches ?? 0;
            if (thicknessIn <= 0)
            {
                ErrorHandler.HandleError("MetricsExtractor.ExtractPartMetrics",
                    $"Part has missing or zero thickness ({thicknessIn} in). Weight and cost calculations will be wrong!",
                    null, ErrorHandler.LogLevel.Warning);
            }

            var pm = new PartMetrics
            {
                MaterialCode = materialCode ?? string.Empty,
                ThicknessIn = thicknessIn,
                BendCount = TryGetBendCount(doc),
                ApproxCutLengthIn = 0, // vNext: compute perimeter+internal cut length
                PierceCount = 0,        // vNext: compute pierces
                WeightCalcMode = rbWeightCalc,
                NestEfficiencyPercent = nestEff,
                BlankLengthIn = blankL,
                BlankWidthIn = blankW,
                Quantity = qty,
                Difficulty = diff
            };

            try
            {
                if (SwMassPropertiesHelper.GetAllMassProperties(doc, out var massKg, out var volM3, out var _, out var __, out var ___))
                {
                    pm.MassKg = massKg;
                    pm.VolumeM3 = volM3;
                }
            }
            catch { }

            return pm;
        }

        private static int TryGetBendCount(IModelDoc2 doc)
        {
            try
            {
                var feat = doc.FirstFeature() as IFeature;
                int count = 0;
                while (feat != null)
                {
                    var t = feat.GetTypeName2() ?? string.Empty;
                    if (t.IndexOf("OneBend", System.StringComparison.OrdinalIgnoreCase) >= 0 ||
                        t.IndexOf("SketchBend", System.StringComparison.OrdinalIgnoreCase) >= 0 ||
                        t.IndexOf("EdgeFlange", System.StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        count++;
                    }
                    feat = feat.GetNextFeature() as IFeature;
                }
                return count;
            }
            catch { return 0; }
        }

        private static double TryParseDouble(string s)
        {
            if (double.TryParse((s ?? string.Empty), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v)) return v;
            if (double.TryParse((s ?? string.Empty), out v)) return v;
            return 0;
        }
    }
}
