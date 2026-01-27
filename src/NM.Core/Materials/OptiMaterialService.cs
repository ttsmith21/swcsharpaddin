using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace NM.Core.Materials
{
    // Resolves OptiMaterial code from Excel object[,] with tolerance-based matching
    public sealed class OptiMaterialService
    {
        private const double THICKNESS_TOLERANCE = 0.008; // inches

        public sealed class OptiMaterialEntry
        {
            public string Code { get; set; }
            public string Material { get; set; }
            public double NominalThickness { get; set; }
            public double MinThickness { get; set; }
            public double MaxThickness { get; set; }
            public string Description { get; set; }
        }

        private readonly Dictionary<string, List<OptiMaterialEntry>> _byMaterial;

        // Expected Excel schema: A=Code, B=Material, C=NominalThickness, D=Min, E=Max, F=Description
        public OptiMaterialService(object[,] excelData)
        {
            var entries = Parse(excelData);
            _byMaterial = entries
                .GroupBy(e => (e.Material ?? string.Empty).ToUpperInvariant())
                .ToDictionary(g => g.Key, g => g.OrderBy(x => x.NominalThickness).ToList(), StringComparer.OrdinalIgnoreCase);
        }

        public string ResolveOptiMaterialCode(double thicknessInches, string material)
        {
            if (string.IsNullOrWhiteSpace(material)) return null;
            if (!_byMaterial.TryGetValue(material.ToUpperInvariant(), out var list) || list == null || list.Count == 0)
                return null;

            // 1) Exact within tolerance
            var exact = list.FirstOrDefault(e => Math.Abs(e.NominalThickness - thicknessInches) <= THICKNESS_TOLERANCE);
            if (exact != null) return exact.Code;

            // 2) Nearest greater
            var greater = list.Where(e => e.NominalThickness > thicknessInches)
                               .OrderBy(e => e.NominalThickness - thicknessInches)
                               .FirstOrDefault();
            if (greater != null) return greater.Code;

            // 3) Max available
            return list.OrderByDescending(e => e.NominalThickness).FirstOrDefault()?.Code;
        }

        public OptiMaterialEntry GetEntry(string optiCode)
        {
            if (string.IsNullOrWhiteSpace(optiCode)) return null;
            foreach (var kv in _byMaterial)
            {
                var match = kv.Value.FirstOrDefault(e => string.Equals(e.Code, optiCode, StringComparison.OrdinalIgnoreCase));
                if (match != null) return match;
            }
            return null;
        }

        private static List<OptiMaterialEntry> Parse(object[,] arr)
        {
            var list = new List<OptiMaterialEntry>();
            if (arr == null) return list;
            int lbR = arr.GetLowerBound(0), ubR = arr.GetUpperBound(0);
            int lbC = arr.GetLowerBound(1), ubC = arr.GetUpperBound(1);

            // Skip header row: assume row 1 has headers ? start at max(2, lbR)
            for (int r = Math.Max(lbR + 1, 2); r <= ubR; r++)
            {
                string code = ToString(arr[r, lbC + 0]);
                if (string.IsNullOrWhiteSpace(code)) continue;
                string material = ToString(arr[r, lbC + 1]);
                double nominal = ToDouble(arr[r, lbC + 2]);
                double minT = ToDouble(arr[r, lbC + 3]);
                double maxT = ToDouble(arr[r, lbC + 4]);
                string desc = (lbC + 5 <= ubC) ? ToString(arr[r, lbC + 5]) : string.Empty;

                list.Add(new OptiMaterialEntry
                {
                    Code = code,
                    Material = material,
                    NominalThickness = nominal,
                    MinThickness = minT,
                    MaxThickness = maxT,
                    Description = desc
                });
            }
            return list;
        }

        private static string ToString(object v) => v?.ToString()?.Trim() ?? string.Empty;
        private static double ToDouble(object v)
        {
            if (v is double d) return d;
            if (v == null) return 0.0;
            if (double.TryParse(v.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var x)) return x;
            if (double.TryParse(v.ToString(), out x)) return x;
            return 0.0;
        }
    }
}
