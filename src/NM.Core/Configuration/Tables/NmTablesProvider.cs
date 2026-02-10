using System.Collections.Generic;

namespace NM.Core.Config.Tables
{
    /// <summary>
    /// Helper methods for looking up values in the NmTables data.
    /// Wraps the raw POCO data with the same fuzzy-match algorithms used by
    /// the legacy static classes (StaticLaserSpeedProvider, TubeCuttingParameterService, etc.).
    /// </summary>
    public static class NmTablesProvider
    {
        /// <summary>
        /// Look up laser speed for a given thickness and material code.
        /// Uses the same algorithm as StaticLaserSpeedProvider.
        /// </summary>
        public static (double feedIpm, double pierceSec) GetLaserSpeed(
            NmTables tables, double thicknessIn, string materialCode)
        {
            if (tables == null) return (0, 0);

            var entries = GetLaserEntriesForMaterial(tables.LaserSpeeds, materialCode);
            if (entries == null || entries.Count == 0) return (0, 0);

            double tolerance = tables.LaserSpeeds.ThicknessToleranceIn;
            double threshold = thicknessIn - tolerance;

            foreach (var e in entries)
            {
                if (e.ThicknessIn >= threshold)
                    return (e.FeedRateIpm, e.PierceSeconds);
            }

            // Fallback: thickest entry
            var last = entries[entries.Count - 1];
            return (last.FeedRateIpm, last.PierceSeconds);
        }

        private static List<LaserSpeedEntry> GetLaserEntriesForMaterial(LaserSpeedTable table, string materialCode)
        {
            var m = (materialCode ?? string.Empty).ToUpperInvariant();

            // 1. Check per-material override first
            if (table.ByMaterial != null && table.ByMaterial.Count > 0)
            {
                if (table.ByMaterial.TryGetValue(m, out var specific) && specific.Count > 0)
                    return specific;
            }

            // 2. Fall back to group tables
            if (m.Contains("A36") || m == "CS" || m.Contains("1018") || m.Contains("1020") || m.Contains("1045"))
                return table.CarbonSteel;

            if (m.Contains("6061") || m.Contains("5052") || m.Contains("3003") || m.Contains("5083"))
                return table.Aluminum;
            if (m == "AL" || m.StartsWith("AL-") || m.EndsWith("-AL"))
                return table.Aluminum;

            return table.StainlessSteel;
        }

        /// <summary>
        /// Look up tube cutting parameters for a given material and wall thickness.
        /// </summary>
        public static (double cutSpeedIpm, double pierceSec, double kerfIn) GetTubeCuttingParams(
            NmTables tables, string materialCategory, double wallIn)
        {
            if (tables == null) return (0, 0, 0.02);

            var material = GetTubeMaterial(tables.TubeCutting, materialCategory);
            if (material == null) return (0, 0, 0.02);

            double feedRate = LookupThreshold(material.FeedRates, wallIn);
            double pierceTime = LookupPierceTime(material.PierceTimes, wallIn);
            double cutSpeed = feedRate * material.FeedMultiplier;

            return (cutSpeed, pierceTime, material.KerfIn);
        }

        private static TubeCuttingMaterial GetTubeMaterial(TubeCuttingTable table, string category)
        {
            switch ((category ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "stainlesssteel": return table.StainlessSteel;
                case "aluminum": return table.Aluminum;
                default: return table.CarbonSteel; // includes "carbonsteel" and unknown
            }
        }

        private static double LookupThreshold(List<TubeFeedRateEntry> entries, double wallIn)
        {
            if (entries == null || entries.Count == 0) return 0;
            foreach (var e in entries)
            {
                if (wallIn <= e.MaxThicknessIn)
                    return e.FeedRateIpm;
            }
            return entries[entries.Count - 1].FeedRateIpm;
        }

        private static double LookupPierceTime(List<TubePierceTimeEntry> entries, double wallIn)
        {
            if (entries == null || entries.Count == 0) return 0;
            foreach (var e in entries)
            {
                if (wallIn <= e.MaxThicknessIn)
                    return e.PierceSeconds;
            }
            return entries[entries.Count - 1].PierceSeconds;
        }
    }
}
