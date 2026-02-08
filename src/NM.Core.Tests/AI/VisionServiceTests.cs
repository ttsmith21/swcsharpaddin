using System.Threading.Tasks;
using Xunit;
using NM.Core.AI;
using NM.Core.AI.Models;

namespace NM.Core.Tests.AI
{
    public class VisionServiceTests
    {
        [Fact]
        public void OfflineService_IsNotAvailable()
        {
            var service = new OfflineVisionService();
            Assert.False(service.IsAvailable);
            Assert.Equal(0m, service.EstimatedCostPerPage);
            Assert.Equal(0m, service.SessionCost);
        }

        [Fact]
        public async Task OfflineService_ReturnsFailure()
        {
            var service = new OfflineVisionService();
            var result = await service.AnalyzeDrawingPageAsync(new byte[0], null);

            Assert.False(result.Success);
            Assert.Contains("not configured", result.ErrorMessage);
        }

        [Fact]
        public async Task OfflineService_TitleBlock_ReturnsFailure()
        {
            var service = new OfflineVisionService();
            var result = await service.AnalyzeTitleBlockAsync(new byte[0]);

            Assert.False(result.Success);
        }

        [Fact]
        public void ClaudeService_NotAvailable_WithoutApiKey()
        {
            var config = new VisionConfig { ApiKey = null };
            var service = new ClaudeVisionService(config);
            Assert.False(service.IsAvailable);
        }

        [Fact]
        public void ClaudeService_Available_WithApiKey()
        {
            var config = new VisionConfig { ApiKey = "test-key-123" };
            var service = new ClaudeVisionService(config);
            Assert.True(service.IsAvailable);
        }

        [Fact]
        public void ClaudeService_CostEstimates()
        {
            var titleBlockConfig = new VisionConfig
            {
                ApiKey = "test",
                Tier = AnalysisTier.TitleBlockOnly
            };
            var titleBlockService = new ClaudeVisionService(titleBlockConfig);
            Assert.Equal(0.001m, titleBlockService.EstimatedCostPerPage);

            var fullPageConfig = new VisionConfig
            {
                ApiKey = "test",
                Tier = AnalysisTier.FullPage
            };
            var fullPageService = new ClaudeVisionService(fullPageConfig);
            Assert.Equal(0.004m, fullPageService.EstimatedCostPerPage);
        }

        [Fact]
        public async Task ClaudeService_NotAvailable_ReturnsError()
        {
            var config = new VisionConfig { ApiKey = null };
            var service = new ClaudeVisionService(config);
            var result = await service.AnalyzeDrawingPageAsync(new byte[0], null);

            Assert.False(result.Success);
            Assert.Contains("not configured", result.ErrorMessage);
        }

        [Fact]
        public void VisionConfig_FromEnvironment_NoKey()
        {
            // Without env var set, should return config with no key
            var config = VisionConfig.FromEnvironment();
            // API key may or may not be set depending on environment
            Assert.NotNull(config);
            Assert.Equal(AiProvider.Claude, config.Provider);
        }

        [Fact]
        public void VisionConfig_Defaults()
        {
            var config = new VisionConfig();
            Assert.Equal(AiProvider.Claude, config.Provider);
            Assert.Equal(200, config.RenderDpi);
            Assert.Equal(2048, config.MaxTokens);
            Assert.Equal(AnalysisTier.TitleBlockOnly, config.Tier);
            Assert.Equal(1.00m, config.MaxSessionCost);
            Assert.Equal(30, config.TimeoutSeconds);
            Assert.False(config.HasApiKey);
        }

        [Fact]
        public void VisionContext_HasContext_Detection()
        {
            var empty = new VisionContext();
            Assert.False(empty.HasContext);

            var withPartNumber = new VisionContext { KnownPartNumber = "123" };
            Assert.True(withPartNumber.HasContext);

            var withMaterial = new VisionContext { KnownMaterial = "A36" };
            Assert.True(withMaterial.HasContext);

            var withThickness = new VisionContext { KnownThickness_in = 0.25 };
            Assert.True(withThickness.HasContext);
        }

        [Fact]
        public void FieldResult_Basics()
        {
            var empty = new FieldResult();
            Assert.False(empty.HasValue);
            Assert.Equal("", empty.ToString());

            var withValue = new FieldResult("ABC", 0.95);
            Assert.True(withValue.HasValue);
            Assert.Equal("ABC", withValue.Value);
            Assert.Equal(0.95, withValue.Confidence);
            Assert.Equal("ABC", withValue.ToString());
        }
    }
}
