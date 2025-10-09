using System;
using System.Collections.Generic;

namespace NM.Core.Materials
{
    public static class MaterialDensityDatabase
    {
        private static readonly Dictionary<string, double> Densities = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
        {
            // Stainless Steel (lb/in^3)
            {"304L", 0.289},
            {"316L", 0.289},
            {"309", 0.287},
            {"2205", 0.283},

            // Carbon Steel
            {"A36", 0.284},
            {"ALNZD", 0.284},
            {"A572", 0.284},

            // Aluminum
            {"6061", 0.098},
            {"5052", 0.097},
            {"3003", 0.099}
        };

        public static double GetDensityLbPerIn3(string material)
        {
            if (string.IsNullOrWhiteSpace(material)) return NM.Core.Manufacturing.Rates.DensityA36;
            if (Densities.TryGetValue(material.ToUpperInvariant(), out var d)) return d;
            return NM.Core.Manufacturing.Rates.GetDensityLbPerIn3(material);
        }

        public static string GetMaterialCategory(string material)
        {
            if (string.IsNullOrWhiteSpace(material)) return "Unknown";
            var u = material.ToUpperInvariant();
            if (u.Contains("304") || u.Contains("316") || u.Contains("309") || u.Contains("2205")) return "StainlessSteel";
            if (u.Contains("606") || u.Contains("505") || u.Contains("3003") || u.StartsWith("AL")) return "Aluminum";
            return "CarbonSteel";
        }
    }
}
