using System;

namespace NM.Core.Manufacturing
{
    public static class MaterialValidator
    {
        // Thickness limits (inches) per material
        public static bool ValidateThickness(string materialCode, double thicknessIn)
        {
            if (string.IsNullOrWhiteSpace(materialCode)) return true; // unknown => skip
            var cat = Classify(materialCode);
            switch (cat)
            {
                case MaterialCategory.Stainless:
                    return thicknessIn >= 0.010 && thicknessIn <= 0.500;
                case MaterialCategory.Carbon:
                    return thicknessIn >= 0.015 && thicknessIn <= 0.750;
                case MaterialCategory.Aluminum:
                    return thicknessIn >= 0.020 && thicknessIn <= 0.250;
                default:
                    return true;
            }
        }

        // Get max bend length (inches) from Bend sheet. Returns 0 if not found
        // bendSheet: values with headers row 1, data from row 2; thickness in col 3; SSN1=5, CS=6, AL=7
        public static double GetMaxBendLengthIn(object[,] bendSheet, string materialCode, double thicknessIn)
        {
            if (bendSheet == null) return 0.0;
            int rows = bendSheet.GetLength(0);
            int cols = bendSheet.GetLength(1);
            int matCol = GetBendMaterialColumn(materialCode);
            if (matCol <= 0 || matCol > cols) return 0.0;

            // Match method: first row where thickness >= (target - 0.005)
            double target = thicknessIn - 0.005;
            for (int r = 2; r <= rows; r++)
            {
                double rowThk = ToDouble(bendSheet[r, 3]);
                if (rowThk >= target)
                {
                    return ToDouble(bendSheet[r, matCol]);
                }
            }
            return 0.0;
        }

        public static bool ValidateTonnage(object[,] bendSheet, string materialCode, double thicknessIn, double longestBendIn, out double maxLen)
        {
            maxLen = GetMaxBendLengthIn(bendSheet, materialCode, thicknessIn);
            if (maxLen <= 0) return true; // no data => skip
            return longestBendIn <= maxLen + 1e-6;
        }

        private static int GetBendMaterialColumn(string materialCode)
        {
            var cat = Classify(materialCode);
            switch (cat)
            {
                case MaterialCategory.Stainless: return 5; // SSN1
                case MaterialCategory.Carbon: return 6; // CS
                case MaterialCategory.Aluminum: return 7; // AL
                default: return 0;
            }
        }

        private enum MaterialCategory { Unknown, Stainless, Carbon, Aluminum }

        private static MaterialCategory Classify(string code)
        {
            try
            {
                var u = (code ?? string.Empty).ToUpperInvariant();
                if (u.Contains("304") || u.Contains("316") || u.Contains("309") || u.Contains("2205")) return MaterialCategory.Stainless;
                if (u.Contains("A36") || u.Contains("ALNZD") || u.Contains("CS")) return MaterialCategory.Carbon;
                if (u.Contains("6061") || u.Contains("5052") || u.Contains("AL")) return MaterialCategory.Aluminum;
            }
            catch { }
            return MaterialCategory.Unknown;
        }

        private static double ToDouble(object v)
        {
            try
            {
                if (v == null) return 0.0;
                if (v is double d) return d;
                if (double.TryParse(v.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var x)) return x;
                if (double.TryParse(v.ToString(), out x)) return x;
                return 0.0;
            }
            catch { return 0.0; }
        }
    }
}
