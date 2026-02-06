using System;
using System.Collections.Generic;

namespace NM.Core.Manufacturing.Laser
{
    /// <summary>
    /// Hardcoded laser speeds for common 304 SS thicknesses with material multipliers.
    /// Fallback when Excel-backed provider is unavailable.
    /// </summary>
    public sealed class StaticLaserSpeedProvider : ILaserSpeedProvider
    {
        private const double THICKNESS_TOLERANCE = 0.005; // inches

        // Base speeds for 304 SS (feed IPM, pierce seconds)
        private static readonly List<SpeedEntry> _entries = new List<SpeedEntry>
        {
            new SpeedEntry(0.048,  200, 0.3),  // 18ga
            new SpeedEntry(0.060,  170, 0.4),  // 16ga
            new SpeedEntry(0.075,  140, 0.5),  // 14ga
            new SpeedEntry(0.105,  100, 0.8),  // 12ga
            new SpeedEntry(0.120,   80, 1.0),  // 11ga
            new SpeedEntry(0.134,   65, 1.2),  // 10ga
            new SpeedEntry(0.1875,  40, 2.0),  // 3/16
            new SpeedEntry(0.250,   25, 3.0),  // 1/4
            new SpeedEntry(0.3125,  18, 4.0),  // 5/16
            new SpeedEntry(0.375,   12, 5.0),  // 3/8
            new SpeedEntry(0.500,    6, 8.0),  // 1/2
        };

        public LaserSpeed GetSpeed(double thicknessIn, string materialCode)
        {
            double multiplier = GetMaterialMultiplier(materialCode);
            if (_entries.Count == 0) return default;

            // 1) exact match within tolerance
            foreach (var e in _entries)
            {
                if (Math.Abs(e.ThicknessIn - thicknessIn) <= THICKNESS_TOLERANCE)
                    return Build(e, multiplier);
            }

            // 2) nearest greater thickness
            foreach (var e in _entries)
            {
                if (e.ThicknessIn > thicknessIn)
                    return Build(e, multiplier);
            }

            // 3) maximum thickness row
            return Build(_entries[_entries.Count - 1], multiplier);
        }

        private static double GetMaterialMultiplier(string material)
        {
            var m = (material ?? string.Empty).ToUpperInvariant();
            if (m.Contains("309") || m.Contains("2205")) return 0.8;
            if (m.Contains("A36") || m.Contains("CS") || m.Contains("1018")) return 1.3;
            if (m.Contains("6061") || m.Contains("5052") || m.Contains("AL")) return 1.5;
            return 1.0; // 304/316 SS baseline
        }

        private static LaserSpeed Build(SpeedEntry e, double feedMultiplier)
        {
            return new LaserSpeed
            {
                FeedRateIpm = e.BaseFeedIpm * feedMultiplier,
                PierceSeconds = e.PierceSeconds
            };
        }

        private readonly struct SpeedEntry
        {
            public readonly double ThicknessIn;
            public readonly double BaseFeedIpm;
            public readonly double PierceSeconds;

            public SpeedEntry(double thickness, double feed, double pierce)
            {
                ThicknessIn = thickness;
                BaseFeedIpm = feed;
                PierceSeconds = pierce;
            }
        }
    }
}
