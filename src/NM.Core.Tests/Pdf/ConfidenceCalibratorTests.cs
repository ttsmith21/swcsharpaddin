using Xunit;
using NM.Core.Pdf;

namespace NM.Core.Tests.Pdf
{
    public class ConfidenceCalibratorTests
    {
        // --- CrossValidateConfidence tests ---

        [Fact]
        public void BothAgree_BoostsConfidence()
        {
            double result = ConfidenceCalibrator.CrossValidateConfidence(
                textConfidence: 0.8, textFound: true,
                visionConfidence: 0.75, visionFound: true);

            // Should boost above the max of the two inputs (0.8)
            Assert.True(result > 0.8, $"Expected > 0.8 but got {result}");
        }

        [Fact]
        public void OnlyTextFound_ReducesConfidence()
        {
            double result = ConfidenceCalibrator.CrossValidateConfidence(
                textConfidence: 0.8, textFound: true,
                visionConfidence: 0, visionFound: false);

            Assert.Equal(0.8 * 0.85, result, 5);
        }

        [Fact]
        public void OnlyVisionFound_ReducesConfidence()
        {
            double result = ConfidenceCalibrator.CrossValidateConfidence(
                textConfidence: 0, textFound: false,
                visionConfidence: 0.9, visionFound: true);

            Assert.Equal(0.9 * 0.85, result, 5);
        }

        [Fact]
        public void NeitherFound_ReturnsZero()
        {
            double result = ConfidenceCalibrator.CrossValidateConfidence(
                textConfidence: 0, textFound: false,
                visionConfidence: 0, visionFound: false);

            Assert.Equal(0, result);
        }

        [Fact]
        public void BoostCapsAtOne()
        {
            double result = ConfidenceCalibrator.CrossValidateConfidence(
                textConfidence: 0.95, textFound: true,
                visionConfidence: 0.95, visionFound: true);

            // 0.95 * 1.15 = 1.0925, should be capped at 1.0
            Assert.Equal(1.0, result);
        }

        // --- CheckCoverageDensity tests ---

        [Fact]
        public void MultiPageFewNotes_IsSuspicious()
        {
            var result = ConfidenceCalibrator.CheckCoverageDensity(
                pageCount: 4, noteCount: 1, gdtCount: 0, hasTolerances: false);

            Assert.True(result.Suspicious);
            Assert.Contains("4 pages", result.Reason);
        }

        [Fact]
        public void MultiPageZeroEverything_IsSuspicious()
        {
            var result = ConfidenceCalibrator.CheckCoverageDensity(
                pageCount: 3, noteCount: 0, gdtCount: 0, hasTolerances: false);

            Assert.True(result.Suspicious);
        }

        [Fact]
        public void SinglePageNoNotes_NotSuspicious()
        {
            var result = ConfidenceCalibrator.CheckCoverageDensity(
                pageCount: 1, noteCount: 0, gdtCount: 0, hasTolerances: false);

            Assert.False(result.Suspicious);
        }

        [Fact]
        public void TitleBlockButZeroNotes_IsSuspicious()
        {
            var result = ConfidenceCalibrator.CheckCoverageDensity(
                pageCount: 1, noteCount: 0, gdtCount: 0,
                hasTolerances: false, hasTitleBlock: true);

            Assert.True(result.Suspicious);
            Assert.Contains("Title block", result.Reason);
        }

        [Fact]
        public void HealthyExtraction_NotSuspicious()
        {
            var result = ConfidenceCalibrator.CheckCoverageDensity(
                pageCount: 2, noteCount: 3, gdtCount: 1, hasTolerances: true);

            Assert.False(result.Suspicious);
        }
    }
}
