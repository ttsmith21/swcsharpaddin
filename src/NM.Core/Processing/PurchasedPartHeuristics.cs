using System;
using System.Collections.Generic;

namespace NM.Core.Processing
{
    /// <summary>
    /// Heuristic analysis to detect parts that are likely purchased (fasteners, fittings, etc.)
    /// based on geometry metrics. Returns a confidence score and reason string for UI display.
    /// This is a hint only - the user still clicks the classification button.
    /// </summary>
    public static class PurchasedPartHeuristics
    {
        public sealed class HeuristicInput
        {
            public double MassKg { get; set; }
            public int FaceCount { get; set; }
            public int EdgeCount { get; set; }
            public double BBoxMaxDimM { get; set; }
            public string FileName { get; set; }
        }

        public sealed class HeuristicResult
        {
            public bool LikelyPurchased { get; set; }
            public double Confidence { get; set; }  // 0.0 to 1.0
            public string Reason { get; set; }
        }

        // Threshold: score >= this value => LikelyPurchased
        private const double PURCHASED_THRESHOLD = 0.5;

        public static HeuristicResult Analyze(HeuristicInput input)
        {
            if (input == null)
                return new HeuristicResult { Reason = string.Empty };

            double score = 0;
            var reasons = new List<string>();

            // Heuristic 1: Very small mass (< 50g)
            if (input.MassKg > 0 && input.MassKg < 0.050)
            {
                score += 0.3;
                reasons.Add($"Low mass ({input.MassKg * 1000:F1}g)");
            }

            // Heuristic 2: High face count (> 50 suggests complex geometry like threads/knurling)
            if (input.FaceCount > 50)
            {
                score += 0.2;
                reasons.Add($"High face count ({input.FaceCount})");
            }

            // Heuristic 3: High edge/face ratio (> 4:1 suggests complex topology)
            if (input.FaceCount > 0 && input.EdgeCount > input.FaceCount * 4)
            {
                score += 0.15;
                reasons.Add($"High edge/face ratio ({input.EdgeCount}/{input.FaceCount})");
            }

            // Heuristic 4: Small bounding box (all dims < 50mm = 0.05m)
            if (input.BBoxMaxDimM > 0 && input.BBoxMaxDimM < 0.05)
            {
                score += 0.15;
                reasons.Add($"Small part ({input.BBoxMaxDimM * 1000:F1}mm max)");
            }

            // Heuristic 5: Filename contains purchased-part keywords
            if (ContainsPurchasedKeyword(input.FileName))
            {
                score += 0.3;
                reasons.Add("Filename suggests purchased");
            }

            return new HeuristicResult
            {
                LikelyPurchased = score >= PURCHASED_THRESHOLD,
                Confidence = Math.Min(score, 1.0),
                Reason = reasons.Count > 0 ? string.Join("; ", reasons) : string.Empty
            };
        }

        private static readonly string[] PurchasedKeywords =
        {
            "bolt", "screw", "nut", "washer", "fastener",
            "fitting", "bushing", "bearing", "motor",
            "gearbox", "actuator", "valve", "swagelock",
            "swagelok", "connector", "coupling", "reducer",
            "elbow", "tee", "nipple", "plug", "cap",
            "o-ring", "oring", "gasket", "seal"
        };

        private static bool ContainsPurchasedKeyword(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return false;
            var lower = fileName.ToLowerInvariant();
            foreach (var kw in PurchasedKeywords)
            {
                if (lower.Contains(kw)) return true;
            }
            return false;
        }
    }
}
