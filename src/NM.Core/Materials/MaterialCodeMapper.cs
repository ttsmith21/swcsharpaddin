using System;
using System.Collections.Generic;

namespace NM.Core.Materials
{
    /// <summary>
    /// Maps SolidWorks material database names to short codes used in descriptions and ERP.
    /// E.g., "AISI 304" -> "304L", "Plain Carbon Steel" -> "CS"
    /// </summary>
    public static class MaterialCodeMapper
    {
        private static readonly Dictionary<string, string> _map =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // Stainless steels (SW DB name -> short code)
            { "AISI 304", "304L" },
            { "304 Stainless Steel", "304L" },
            { "304 Stainless Steel (SS)", "304L" },
            { "AISI 316", "316L" },
            { "316 Stainless Steel", "316L" },
            { "316 Stainless Steel (SS)", "316L" },
            { "AISI 309", "309" },
            { "AISI 321", "321" },
            { "AISI 430", "430" },
            { "AISI 201", "201" },

            // Carbon steels
            { "Plain Carbon Steel", "CS" },
            { "AISI 1018 Steel", "1018" },
            { "1018 Steel", "1018" },
            { "AISI 1020 Steel", "1020" },
            { "AISI 1045 Steel", "1045" },
            { "A36 Steel", "A36" },
            { "ASTM A36 Steel", "A36" },

            // Aluminum
            { "6061 Alloy", "6061" },
            { "6061-T6 (SS)", "6061" },
            { "5052 Alloy", "5052" },
            { "5052-H32", "5052" },
            { "3003 Alloy", "3003" },
            { "5083 Alloy", "5083" },

            // Pass-through (already short codes — all 18 form materials)
            { "304L", "304L" },
            { "316L", "316L" },
            { "309", "309" },
            { "310", "310" },
            { "321", "321" },
            { "330", "330" },
            { "409", "409" },
            { "430", "430" },
            { "2205", "2205" },
            { "2507", "2507" },
            { "C22", "C22" },
            { "C276", "C276" },
            { "AL6XN", "AL6XN" },
            { "ALLOY31", "ALLOY31" },
            { "CS", "CS" },
            { "A36", "A36" },
            { "ALNZD", "ALNZD" },
            { "1018", "1018" },
            { "6061", "6061" },
            { "5052", "5052" },
        };

        /// <summary>
        /// Convert a SolidWorks material name to a short code for descriptions.
        /// Returns the original material (uppercased) if no mapping exists.
        /// </summary>
        public static string ToShortCode(string swMaterialName)
        {
            if (string.IsNullOrWhiteSpace(swMaterialName))
                return null;

            string trimmed = swMaterialName.Trim();

            if (_map.TryGetValue(trimmed, out string code))
                return code;

            // Fallback: strip common prefixes
            string upper = trimmed.ToUpperInvariant();
            if (upper.StartsWith("AISI "))
                return upper.Substring(5).Trim();

            return upper;
        }
    }
}
