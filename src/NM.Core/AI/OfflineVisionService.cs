using System.Threading.Tasks;
using NM.Core.AI.Models;

namespace NM.Core.AI
{
    /// <summary>
    /// Offline fallback when no AI API key is configured.
    /// Always returns empty results, allowing the pipeline to fall back
    /// to PdfPig text extraction (Phase 1).
    /// </summary>
    public sealed class OfflineVisionService : IDrawingVisionService
    {
        public bool IsAvailable => false;

        public decimal EstimatedCostPerPage => 0m;

        public decimal SessionCost => 0m;

        public Task<VisionAnalysisResult> AnalyzeDrawingPageAsync(byte[] imageBytes, VisionContext context)
        {
            return Task.FromResult(new VisionAnalysisResult
            {
                Success = false,
                ErrorMessage = "AI Vision not configured. Using offline text extraction only. " +
                               "Set NM_AI_API_KEY environment variable to enable AI analysis."
            });
        }

        public Task<VisionAnalysisResult> AnalyzeTitleBlockAsync(byte[] imageBytes)
        {
            return Task.FromResult(new VisionAnalysisResult
            {
                Success = false,
                ErrorMessage = "AI Vision not configured. Using offline text extraction only."
            });
        }
    }
}
