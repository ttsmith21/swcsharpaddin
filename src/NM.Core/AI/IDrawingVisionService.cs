using System.Threading.Tasks;
using NM.Core.AI.Models;

namespace NM.Core.AI
{
    /// <summary>
    /// Interface for AI-powered drawing analysis.
    /// Implementations can use Claude, GPT-4o, or offline heuristics.
    /// </summary>
    public interface IDrawingVisionService
    {
        /// <summary>
        /// Analyzes a PDF drawing page image using AI vision.
        /// Returns null if the service is unavailable.
        /// </summary>
        /// <param name="imageBytes">PNG image bytes of the drawing page.</param>
        /// <param name="context">Optional context to guide the analysis (e.g., known part info).</param>
        Task<VisionAnalysisResult> AnalyzeDrawingPageAsync(byte[] imageBytes, VisionContext context);

        /// <summary>
        /// Analyzes only the title block region of a drawing page.
        /// Cheaper than full-page analysis (smaller image).
        /// </summary>
        /// <param name="imageBytes">PNG image bytes of the title block region.</param>
        Task<VisionAnalysisResult> AnalyzeTitleBlockAsync(byte[] imageBytes);

        /// <summary>
        /// True if the service is configured and ready to accept requests.
        /// </summary>
        bool IsAvailable { get; }

        /// <summary>
        /// Estimated cost in USD for analyzing one full drawing page.
        /// </summary>
        decimal EstimatedCostPerPage { get; }

        /// <summary>
        /// Cumulative cost in USD for this session.
        /// </summary>
        decimal SessionCost { get; }
    }
}
