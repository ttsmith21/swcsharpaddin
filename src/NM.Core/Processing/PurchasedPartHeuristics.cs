using System;
using System.Collections.Generic;

namespace NM.Core.Processing
{
    /// <summary>
    /// Heuristic analysis to detect parts that are likely purchased (fasteners, fittings, etc.)
    /// based on geometry metrics. Returns a confidence score and reason string for UI display.
    /// This is a hint only - the user still clicks the classification button.
    ///
    /// Score must reach 1.0 to flag as likely purchased.
    /// </summary>
    public static class PurchasedPartHeuristics
    {
        public sealed class HeuristicInput
        {
            public double MassKg { get; set; }
            public int FaceCount { get; set; }
            public int EdgeCount { get; set; }
            public double BBoxMaxDimM { get; set; }
            public double BBoxMinDimM { get; set; }
            public string FileName { get; set; }
        }

        public sealed class HeuristicResult
        {
            public bool LikelyPurchased { get; set; }
            public double Confidence { get; set; }  // 0.0 to 1.0
            public string Reason { get; set; }
        }

        private const double PURCHASED_THRESHOLD = 1.0;

        public static HeuristicResult Analyze(HeuristicInput input)
        {
            if (input == null)
                return new HeuristicResult { Reason = string.Empty };

            double score = 0;
            var reasons = new List<string>();

            // === MASS HEURISTICS ===
            // < 100g: definite flag (score = 1.0 alone)
            if (input.MassKg > 0 && input.MassKg < 0.100)
            {
                score += 1.0;
                reasons.Add($"Low mass ({input.MassKg * 1000:F1}g)");
            }
            // 100g - 500g: suspect range, needs another signal
            else if (input.MassKg >= 0.100 && input.MassKg < 0.500)
            {
                score += 0.50;
                reasons.Add($"Suspect mass ({input.MassKg * 1000:F0}g)");
            }

            // === GEOMETRY HEURISTICS ===
            // High face count (> 50 suggests threads/knurling)
            if (input.FaceCount > 50)
            {
                score += 0.50;
                reasons.Add($"High face count ({input.FaceCount})");
            }

            // High edge/face ratio (> 4:1 suggests complex topology)
            if (input.FaceCount > 0 && input.EdgeCount > input.FaceCount * 4)
            {
                score += 0.35;
                reasons.Add($"High edge/face ratio ({input.EdgeCount}/{input.FaceCount})");
            }

            // Small bounding box (max dim < 50mm)
            if (input.BBoxMaxDimM > 0 && input.BBoxMaxDimM < 0.05)
            {
                score += 0.35;
                reasons.Add($"Small part ({input.BBoxMaxDimM * 1000:F1}mm max)");
            }

            // Compact shape (max/min ratio < 2:1 AND small < 75mm)
            // Cube-like parts are fasteners/fittings, not flat plates
            if (input.BBoxMinDimM > 0 && input.BBoxMaxDimM > 0)
            {
                double ratio = input.BBoxMaxDimM / input.BBoxMinDimM;
                if (ratio < 2.0 && input.BBoxMaxDimM < 0.075)
                {
                    score += 0.50;
                    reasons.Add($"Compact shape (ratio {ratio:F1}:1)");
                }
            }

            // === KEYWORD HEURISTICS ===
            double kwScore = GetKeywordScore(input.FileName);
            if (kwScore > 0)
            {
                score += kwScore;
                reasons.Add(kwScore >= 1.0
                    ? "Filename: definite purchased"
                    : "Filename: suspect purchased");
            }

            return new HeuristicResult
            {
                LikelyPurchased = score >= PURCHASED_THRESHOLD,
                Confidence = Math.Min(score, 1.0),
                Reason = reasons.Count > 0 ? string.Join("; ", reasons) : string.Empty
            };
        }

        /// <summary>
        /// Returns the keyword score for a filename. Checks tiers in order
        /// and returns the highest matching score.
        /// </summary>
        private static double GetKeywordScore(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return 0;
            var lower = fileName.ToLowerInvariant();

            // Tier 1: Definite purchased (1.0) — hardware, fittings, seals, thread sizes
            foreach (var kw in DefiniteKeywords)
            {
                if (lower.Contains(kw)) return 1.0;
            }

            // Tier 2: Suspect (0.50) — could go either way
            foreach (var kw in SuspectKeywords)
            {
                if (lower.Contains(kw)) return 0.50;
            }

            return 0;
        }

        // Tier 1: Score = 1.0 (definite purchased, flags alone)
        private static readonly string[] DefiniteKeywords =
        {
            // --- Hardware ---
            "bolt", "screw", "nut", "washer", "fastener", "rivet",
            "cotter", "dowel",

            // --- Fastener acronyms ---
            "shcs",     // Socket Head Cap Screw
            "hhcs",     // Hex Head Cap Screw
            "fhcs",     // Flat Head Cap Screw
            "bhcs",     // Button Head Cap Screw
            "sss",      // Socket Set Screw
            "pem",      // Self-clinching hardware (brand)

            // --- Standard codes ---
            "din",      // DIN 912, DIN 934, etc.
            "iso ",     // ISO 4762, ISO 7380
            "iso_",
            "ansi",     // ANSI B18.2
            "asme",
            "nas",      // National Aerospace Standard
            "nasm",
            "jis",      // Japanese Industrial Standard

            // --- Thread specifications ---
            "unc", "unf",       // Unified Coarse/Fine
            "npt", "bsp",       // Pipe threads

            // --- Drive types ---
            "torx", "phillips", "slotted", "pozi",

            // --- Bearings & drive ---
            "bearing", "motor", "gearbox", "actuator",

            // --- Fittings ---
            "fitting", "bushing", "connector", "coupling", "reducer",
            "union", "adapter", "elbow", "tee", "nipple", "plug", "cap",

            // --- Valves ---
            "valve", "swagelok", "swagelock",

            // --- Seals ---
            "o-ring", "oring", "gasket", "seal",

            // --- Clips & retaining ---
            "circlip", "e-clip", "retaining ring", "snap ring",

            // --- Thread inserts ---
            "helicoil", "keensert",

            // --- Other definite ---
            "hose", "clamp", "spring", "stud", "zerk", "caster",

            // --- Vendor names ---
            "mcmaster", "carr", "grainger", "fastenal", "misumi",
            "mouser", "digikey", "unistrut", "80/20",

            // --- Bolt grades (fastener-specific) ---
            "grade 8", "grade_8",
            "class 10.9", "class_10.9",
            "class 12.9", "class_12.9",
            "18-8",             // Stainless steel fastener grade

            // --- Metric thread sizes ---
            "m3", "m4", "m5", "m6", "m8", "m10", "m12",
            "m14", "m16", "m20", "m24",

            // --- Imperial thread sizes ---
            "1/4-20", "5/16-18", "3/8-16", "1/2-13",
            "1/4-28", "5/16-24", "3/8-24", "1/2-20",
            "#4-40", "#6-32", "#8-32", "#10-24", "#10-32"
        };

        // Tier 2: Score = 0.50 (suspect, needs geometry/mass signal)
        private static readonly string[] SuspectKeywords =
        {
            "insert", "spacer", "standoff", "pin", "anchor"
        };
    }
}
