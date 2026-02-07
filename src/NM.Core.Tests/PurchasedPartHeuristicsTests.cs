using NM.Core.Processing;
using Xunit;

namespace NM.Core.Tests
{
    public class PurchasedPartHeuristicsTests
    {
        [Fact]
        public void NullInput_ReturnsNotPurchased()
        {
            var result = PurchasedPartHeuristics.Analyze(null);
            Assert.False(result.LikelyPurchased);
            Assert.Equal(0, result.Confidence);
        }

        [Fact]
        public void EmptyInput_ReturnsNotPurchased()
        {
            var input = new PurchasedPartHeuristics.HeuristicInput();
            var result = PurchasedPartHeuristics.Analyze(input);
            Assert.False(result.LikelyPurchased);
        }

        [Fact]
        public void SmallBolt_LikelyPurchased()
        {
            // Small bolt: low mass, high face count, keyword match
            var input = new PurchasedPartHeuristics.HeuristicInput
            {
                MassKg = 0.010,     // 10 grams
                FaceCount = 60,     // threaded bolt has many faces
                EdgeCount = 300,    // high edge count
                FileName = "M10_Bolt.sldprt"
            };
            var result = PurchasedPartHeuristics.Analyze(input);
            Assert.True(result.LikelyPurchased);
            Assert.True(result.Confidence >= 0.5);
            Assert.Contains("Low mass", result.Reason);
            Assert.Contains("Filename suggests purchased", result.Reason);
        }

        [Fact]
        public void NormalBracket_NotPurchased()
        {
            // Normal sheet metal bracket: medium mass, low face count
            var input = new PurchasedPartHeuristics.HeuristicInput
            {
                MassKg = 0.500,     // 500g
                FaceCount = 12,
                EdgeCount = 24,
                FileName = "Bracket_001.sldprt"
            };
            var result = PurchasedPartHeuristics.Analyze(input);
            Assert.False(result.LikelyPurchased);
        }

        [Fact]
        public void SwagelokFitting_KeywordMatch()
        {
            var input = new PurchasedPartHeuristics.HeuristicInput
            {
                MassKg = 0.080,     // 80g
                FaceCount = 40,
                EdgeCount = 80,
                FileName = "D8_Swagelock_T_Fitting.sldprt"
            };
            var result = PurchasedPartHeuristics.Analyze(input);
            // "swagelock" keyword + low mass = score >= 0.5
            Assert.True(result.LikelyPurchased);
            Assert.Contains("Filename suggests purchased", result.Reason);
        }

        [Fact]
        public void SmallPart_WithSmallBBox()
        {
            var input = new PurchasedPartHeuristics.HeuristicInput
            {
                MassKg = 0.005,      // 5g
                FaceCount = 20,
                EdgeCount = 40,
                BBoxMaxDimM = 0.020, // 20mm
                FileName = "Pin_001.sldprt"
            };
            var result = PurchasedPartHeuristics.Analyze(input);
            // Low mass (0.3) + small part (0.15) = 0.45, just under threshold
            // Should NOT be LikelyPurchased without keyword or high face count
            Assert.Equal(0.45, result.Confidence, 2);
        }

        [Fact]
        public void Confidence_NeverExceedsOne()
        {
            // Max everything: low mass + high faces + high edge ratio + small bbox + keyword
            var input = new PurchasedPartHeuristics.HeuristicInput
            {
                MassKg = 0.001,
                FaceCount = 100,
                EdgeCount = 800,
                BBoxMaxDimM = 0.01,
                FileName = "M6_Bolt_Washer.sldprt"
            };
            var result = PurchasedPartHeuristics.Analyze(input);
            Assert.True(result.LikelyPurchased);
            Assert.True(result.Confidence <= 1.0);
        }

        [Fact]
        public void HighFaceCount_ContributesToScore()
        {
            var input = new PurchasedPartHeuristics.HeuristicInput
            {
                MassKg = 0.200,     // not low enough for mass heuristic
                FaceCount = 80,     // high face count
                EdgeCount = 400,    // high edge/face ratio
                FileName = "Custom_Part.sldprt"
            };
            var result = PurchasedPartHeuristics.Analyze(input);
            Assert.Contains("High face count", result.Reason);
            Assert.Contains("High edge/face ratio", result.Reason);
        }
    }
}
