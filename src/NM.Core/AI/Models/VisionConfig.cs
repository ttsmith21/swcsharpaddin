using System;
using System.Collections.Generic;

namespace NM.Core.AI.Models
{
    /// <summary>
    /// Configuration for the AI vision service.
    /// </summary>
    public sealed class VisionConfig
    {
        /// <summary>API key for the AI provider (Claude, OpenAI, etc.).</summary>
        public string ApiKey { get; set; }

        /// <summary>AI provider to use.</summary>
        public AiProvider Provider { get; set; } = AiProvider.Claude;

        /// <summary>Model ID to use (e.g., "claude-sonnet-4-5-20250929").</summary>
        public string ModelId { get; set; } = "claude-sonnet-4-5-20250929";

        /// <summary>DPI for rendering PDF pages to images. Higher = better accuracy, more cost.</summary>
        public int RenderDpi { get; set; } = 200;

        /// <summary>Maximum tokens for the AI response.</summary>
        public int MaxTokens { get; set; } = 2048;

        /// <summary>Analysis tier: controls cost/accuracy tradeoff.</summary>
        public AnalysisTier Tier { get; set; } = AnalysisTier.TitleBlockOnly;

        /// <summary>Maximum cost per session in USD. Service stops when exceeded.</summary>
        public decimal MaxSessionCost { get; set; } = 1.00m;

        /// <summary>Timeout per API call in seconds.</summary>
        public int TimeoutSeconds { get; set; } = 30;

        /// <summary>True if the configuration has a valid API key.</summary>
        public bool HasApiKey => !string.IsNullOrWhiteSpace(ApiKey);

        /// <summary>
        /// Loads config from environment variable (NM_AI_API_KEY) or a config file.
        /// </summary>
        public static VisionConfig FromEnvironment()
        {
            var config = new VisionConfig();

            // Try environment variable first
            string apiKey = Environment.GetEnvironmentVariable("NM_AI_API_KEY");
            if (!string.IsNullOrWhiteSpace(apiKey))
            {
                config.ApiKey = apiKey;
            }

            // Check for provider override
            string provider = Environment.GetEnvironmentVariable("NM_AI_PROVIDER");
            if (!string.IsNullOrWhiteSpace(provider))
            {
                if (Enum.TryParse<AiProvider>(provider, true, out var p))
                    config.Provider = p;
            }

            // Check for model override
            string model = Environment.GetEnvironmentVariable("NM_AI_MODEL");
            if (!string.IsNullOrWhiteSpace(model))
            {
                config.ModelId = model;
            }

            // Check for tier override
            string tier = Environment.GetEnvironmentVariable("NM_AI_TIER");
            if (!string.IsNullOrWhiteSpace(tier))
            {
                if (Enum.TryParse<AnalysisTier>(tier, true, out var t))
                    config.Tier = t;
            }

            return config;
        }
    }

    public enum AiProvider
    {
        Claude,
        Gemini,
        OpenAI,
        Offline
    }

    public enum AnalysisTier
    {
        /// <summary>Text extraction only, no AI. Free.</summary>
        TextOnly,

        /// <summary>AI analyzes title block region only. ~$0.001/page.</summary>
        TitleBlockOnly,

        /// <summary>AI analyzes entire drawing page. ~$0.004/page.</summary>
        FullPage
    }
}
