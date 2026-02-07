using System;
using System.Collections.Generic;

namespace NM.Core.Manufacturing.Laser
{
    /// <summary>
    /// Laser speed/pierce data from Laser2022v4.xlsx, hardcoded for headless/QA mode.
    /// Three separate tables per material category (Stainless Steel, Carbon Steel, Aluminum).
    /// Matching algorithm: VBA parity — find thinnest entry where thickness >= (partThickness - 0.005").
    /// </summary>
    public sealed class StaticLaserSpeedProvider : ILaserSpeedProvider
    {
        private const double THICKNESS_TOLERANCE = 0.005; // inches

        // Stainless Steel (F120 fiber laser) - 28 entries, sorted ascending by thickness
        private static readonly List<SpeedEntry> _ssEntries = new List<SpeedEntry>
        {
            new SpeedEntry(0.001,  1600, 0.01),   // Gauge 28
            new SpeedEntry(0.020,  1600, 0.01),   // Gauge 26
            new SpeedEntry(0.024,  1600, 0.01),   // Gauge 24
            new SpeedEntry(0.030,  1600, 0.01),   // Gauge 22
            new SpeedEntry(0.036,  1600, 0.01),   // Gauge 20
            new SpeedEntry(0.048,  1600, 0.1),    // Gauge 18
            new SpeedEntry(0.060,  1600, 0.1),    // Gauge 16
            new SpeedEntry(0.075,  1570, 0.14),   // Gauge 14
            new SpeedEntry(0.094,  1100, 0.15),   // Gauge 13
            new SpeedEntry(0.105,  820,  0.15),   // Gauge 12
            new SpeedEntry(0.120,  253,  0.0467), // Gauge 11
            new SpeedEntry(0.135,  545,  0.15),   // Gauge 10
            new SpeedEntry(0.149,  370,  0.15),   // Gauge 9
            new SpeedEntry(0.165,  300,  0.2),    // Gauge 8
            new SpeedEntry(0.179,  300,  0.2),    // Gauge 7
            new SpeedEntry(0.188,  480,  0.1),    // 3/16 Plate
            new SpeedEntry(0.250,  320,  0.25),   // 1/4 Plate
            new SpeedEntry(0.3125, 240,  0.25),   // 5/16 Plate
            new SpeedEntry(0.375,  138,  0.5),    // 3/8 Plate
            new SpeedEntry(0.500,  135,  1.2),    // 1/2 Plate
            new SpeedEntry(0.625,  88,   2.0),    // 5/8 Plate
            new SpeedEntry(0.750,  56,   3.0),    // 3/4 Plate
            new SpeedEntry(0.875,  33,   3.0),    // 7/8 Plate
            new SpeedEntry(1.000,  21,   3.0),    // 1" Plate
            new SpeedEntry(1.250,  6,    10.0),   // 1.25" Plate
            new SpeedEntry(1.500,  6,    10.0),   // 1.5" Plate
            new SpeedEntry(2.000,  0.001, 999.0), // 2" Plate (not cuttable)
            new SpeedEntry(3.000,  0.001, 999.0), // 3" Plate (not cuttable)
        };

        // Carbon Steel (F120 fiber laser) - 25 entries, sorted ascending by thickness
        private static readonly List<SpeedEntry> _csEntries = new List<SpeedEntry>
        {
            new SpeedEntry(0.018,  2300, 0.01),   // Gauge 26
            new SpeedEntry(0.024,  2300, 0.01),   // Gauge 24
            new SpeedEntry(0.030,  2300, 0.01),   // Gauge 22
            new SpeedEntry(0.036,  2300, 0.01),   // Gauge 20
            new SpeedEntry(0.048,  2100, 0.01),   // Gauge 18
            new SpeedEntry(0.060,  1800, 0.01),   // Gauge 16
            new SpeedEntry(0.075,  1400, 0.06),   // Gauge 14
            new SpeedEntry(0.094,  1100, 0.06),   // Gauge 13
            new SpeedEntry(0.105,  830,  0.1),    // Gauge 12
            new SpeedEntry(0.120,  810,  0.1),    // Gauge 11
            new SpeedEntry(0.135,  670,  0.1),    // Gauge 10
            new SpeedEntry(0.149,  550,  0.1),    // Gauge 9
            new SpeedEntry(0.165,  420,  0.2),    // Gauge 8
            new SpeedEntry(0.179,  420,  0.2),    // Gauge 7
            new SpeedEntry(0.188,  320,  0.4),    // 3/16 Plate
            new SpeedEntry(0.250,  180,  0.4),    // 1/4 Plate
            new SpeedEntry(0.3125, 100,  0.2),    // 5/16 Plate
            new SpeedEntry(0.375,  95,   1.0),    // 3/8 Plate
            new SpeedEntry(0.500,  60,   3.0),    // 1/2 Plate
            new SpeedEntry(0.625,  65,   5.0),    // 5/8 Plate
            new SpeedEntry(0.750,  55,   8.0),    // 3/4 Plate
            new SpeedEntry(1.000,  35,   10.0),   // 1" Plate
            new SpeedEntry(1.500,  1.66, 80.0),   // 1.5" Plate
            new SpeedEntry(2.000,  1.18, 80.0),   // 2" Plate
            new SpeedEntry(3.000,  0.57, 80.0),   // 3" Plate
        };

        // Aluminum (F120 fiber laser) - 24 entries, sorted ascending by thickness
        private static readonly List<SpeedEntry> _alEntries = new List<SpeedEntry>
        {
            new SpeedEntry(0.024,  3100, 0.01),   // Gauge 24
            new SpeedEntry(0.030,  2700, 0.01),   // Gauge 22
            new SpeedEntry(0.036,  2700, 0.01),   // Gauge 20
            new SpeedEntry(0.048,  2700, 0.01),   // Gauge 18
            new SpeedEntry(0.060,  2700, 0.01),   // Gauge 16
            new SpeedEntry(0.075,  2700, 0.04),   // Gauge 14
            new SpeedEntry(0.094,  2200, 0.04),   // Gauge 13
            new SpeedEntry(0.105,  1700, 0.06),   // Gauge 12
            new SpeedEntry(0.120,  1700, 0.06),   // Gauge 11
            new SpeedEntry(0.135,  1400, 0.06),   // Gauge 10
            new SpeedEntry(0.149,  1100, 0.06),   // Gauge 9
            new SpeedEntry(0.165,  230,  0.07),   // Gauge 8
            new SpeedEntry(0.179,  230,  0.07),   // Gauge 7
            new SpeedEntry(0.188,  430,  0.07),   // 3/16 Plate
            new SpeedEntry(0.250,  280,  0.1),    // 1/4 Plate
            new SpeedEntry(0.3125, 185,  0.25),   // 5/16 Plate
            new SpeedEntry(0.375,  120,  0.6),    // 3/8 Plate
            new SpeedEntry(0.500,  70,   1.1),    // 1/2 Plate
            new SpeedEntry(0.625,  25,   2.5),    // 5/8 Plate
            new SpeedEntry(0.750,  23,   10.0),   // 3/4 Plate
            new SpeedEntry(1.000,  7,    20.0),   // 1" Plate
            new SpeedEntry(1.500,  1.66, 80.0),   // 1.5" Plate
            new SpeedEntry(2.000,  1.18, 80.0),   // 2" Plate
            new SpeedEntry(3.000,  0.57, 80.0),   // 3" Plate
        };

        public LaserSpeed GetSpeed(double thicknessIn, string materialCode)
        {
            var entries = GetEntriesForMaterial(materialCode);
            if (entries.Count == 0) return default;

            // VBA algorithm: find thinnest entry where thickness >= (partThickness - tolerance)
            // Data is sorted ascending; first match is the thinnest qualifying entry.
            double threshold = thicknessIn - THICKNESS_TOLERANCE;
            foreach (var e in entries)
            {
                if (e.ThicknessIn >= threshold)
                {
                    return new LaserSpeed
                    {
                        FeedRateIpm = e.FeedRateIpm,
                        PierceSeconds = e.PierceSeconds
                    };
                }
            }

            // Fallback: thickest available entry (part is thicker than all entries)
            var last = entries[entries.Count - 1];
            return new LaserSpeed
            {
                FeedRateIpm = last.FeedRateIpm,
                PierceSeconds = last.PierceSeconds
            };
        }

        private static List<SpeedEntry> GetEntriesForMaterial(string materialCode)
        {
            var m = (materialCode ?? string.Empty).ToUpperInvariant();

            // Carbon steel
            if (m.Contains("A36") || m == "CS" || m.Contains("1018") || m.Contains("1020") || m.Contains("1045"))
                return _csEntries;

            // Aluminum
            if (m.Contains("6061") || m.Contains("5052") || m.Contains("3003") || m.Contains("5083"))
                return _alEntries;
            // Check "AL" last to avoid matching "304L" which contains "AL" — nope, 304L doesn't contain "AL"
            // Actually "ALNZD" is aluminized steel (carbon steel category), not aluminum
            if (m == "AL" || m.StartsWith("AL-") || m.EndsWith("-AL"))
                return _alEntries;

            // Default: Stainless Steel (304L, 316L, 309, 2205, 321, 430, 201, etc.)
            return _ssEntries;
        }

        private readonly struct SpeedEntry
        {
            public readonly double ThicknessIn;
            public readonly double FeedRateIpm;
            public readonly double PierceSeconds;

            public SpeedEntry(double thickness, double feed, double pierce)
            {
                ThicknessIn = thickness;
                FeedRateIpm = feed;
                PierceSeconds = pierce;
            }
        }
    }
}
