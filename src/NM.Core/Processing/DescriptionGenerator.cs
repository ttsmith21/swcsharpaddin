using System;
using NM.Core.DataModel;
using NM.Core.Materials;

namespace NM.Core.Processing
{
    /// <summary>
    /// Generates VBA-parity part descriptions from PartData.
    /// Format: "{MATERIAL} {SUFFIX}" where suffix depends on classification and shape.
    /// </summary>
    public static class DescriptionGenerator
    {
        /// <summary>
        /// Generate a description string for the given part data.
        /// Returns null if pd is null or Material is empty/whitespace.
        /// </summary>
        public static string Generate(PartData pd)
        {
            if (pd == null) return null;
            if (string.IsNullOrWhiteSpace(pd.Material)) return null;

            // Priority: user-form MaterialCode in Extra > MaterialCodeMapper > raw material
            string material;
            string materialCode;
            if (pd.Extra.TryGetValue("MaterialCode", out materialCode) &&
                !string.IsNullOrWhiteSpace(materialCode))
            {
                material = materialCode.Trim().ToUpperInvariant();
            }
            else
            {
                material = MaterialCodeMapper.ToShortCode(pd.Material);
            }

            if (string.IsNullOrWhiteSpace(material)) return null;

            string suffix = GetSuffix(pd);

            if (string.IsNullOrEmpty(suffix))
                return material;

            return material + " " + suffix;
        }

        private static string GetSuffix(PartData pd)
        {
            if (pd.Classification == PartType.SheetMetal || (pd.Sheet != null && pd.Sheet.IsSheetMetal))
                return GetSheetMetalSuffix(pd);

            if (pd.Classification == PartType.Tube || (pd.Tube != null && pd.Tube.IsTube))
                return GetTubeSuffix(pd);

            // Generic, Assembly, Purchased, Unknown - no suffix
            return null;
        }

        private static string GetSheetMetalSuffix(PartData pd)
        {
            // Roll forming takes precedence over bends
            // Check F325 price (if costs already computed) OR max bend radius > 2" (roll forming threshold)
            if (pd.Cost != null && pd.Cost.F325_Price > 0)
                return "ROLL";
            if (pd.Extra != null)
            {
                string radiusStr;
                if (pd.Extra.TryGetValue("MaxBendRadiusIn", out radiusStr))
                {
                    double radius;
                    if (double.TryParse(radiusStr, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out radius) && radius > 2.0)
                        return "ROLL";
                }
            }

            if (pd.Sheet != null && pd.Sheet.BendCount > 0)
                return "BENT";

            // Flat sheet - no bends, no roll forming
            return "PLATE";
        }

        private static string GetTubeSuffix(PartData pd)
        {
            if (pd.Tube == null)
                return "PIPE"; // default to round

            string shape = (pd.Tube.TubeShape ?? "").Trim();

            if (shape.Equals("Square", StringComparison.OrdinalIgnoreCase))
                return "SQ TUBE";

            if (shape.Equals("Rectangle", StringComparison.OrdinalIgnoreCase))
                return "RECT TUBE";

            if (shape.Equals("Angle", StringComparison.OrdinalIgnoreCase))
                return "ANGLE";

            if (shape.Equals("Channel", StringComparison.OrdinalIgnoreCase))
                return "CHANNEL";

            if (shape.Equals("Round Bar", StringComparison.OrdinalIgnoreCase))
                return "ROUND";

            // Default for Round or unrecognized shapes
            return "PIPE";
        }
    }
}
