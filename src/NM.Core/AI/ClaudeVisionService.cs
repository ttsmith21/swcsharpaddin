using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NM.Core.AI.Models;

namespace NM.Core.AI
{
    /// <summary>
    /// AI drawing analysis using the Anthropic Claude API with vision capabilities.
    /// Uses raw HttpWebRequest (no external SDK dependency) for .NET Framework 4.8 compatibility.
    /// </summary>
    public sealed class ClaudeVisionService : IDrawingVisionService
    {
        private const string ApiUrl = "https://api.anthropic.com/v1/messages";
        private const string ApiVersion = "2023-06-01";

        private readonly VisionConfig _config;
        private readonly VisionResponseParser _parser;
        private decimal _sessionCost;

        public ClaudeVisionService(VisionConfig config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
            _parser = new VisionResponseParser();
        }

        public bool IsAvailable => _config.HasApiKey && _sessionCost < _config.MaxSessionCost;

        public decimal EstimatedCostPerPage
        {
            get
            {
                switch (_config.Tier)
                {
                    case AnalysisTier.TitleBlockOnly: return 0.001m;
                    case AnalysisTier.FullPage: return 0.004m;
                    default: return 0m;
                }
            }
        }

        public decimal SessionCost => _sessionCost;

        public async Task<VisionAnalysisResult> AnalyzeDrawingPageAsync(byte[] imageBytes, VisionContext context)
        {
            if (!IsAvailable)
            {
                return new VisionAnalysisResult
                {
                    Success = false,
                    ErrorMessage = _sessionCost >= _config.MaxSessionCost
                        ? $"Session cost limit reached (${_sessionCost:F3} >= ${_config.MaxSessionCost:F2})"
                        : "Claude Vision service not configured (missing API key)"
                };
            }

            string prompt = (context != null && context.HasContext)
                ? VisionPrompts.GetFullPagePromptWithContext(context)
                : VisionPrompts.GetFullPagePrompt();

            string systemPrompt = VisionPrompts.GetSystemPrompt();

            return await CallClaudeApiAsync(imageBytes, prompt, systemPrompt, fullPage: true);
        }

        public async Task<VisionAnalysisResult> AnalyzeTitleBlockAsync(byte[] imageBytes)
        {
            if (!IsAvailable)
            {
                return new VisionAnalysisResult
                {
                    Success = false,
                    ErrorMessage = "Claude Vision service not available"
                };
            }

            string prompt = VisionPrompts.GetTitleBlockPrompt();
            string systemPrompt = VisionPrompts.GetSystemPrompt();

            return await CallClaudeApiAsync(imageBytes, prompt, systemPrompt, fullPage: false);
        }

        private async Task<VisionAnalysisResult> CallClaudeApiAsync(
            byte[] imageBytes, string userPrompt, string systemPrompt, bool fullPage)
        {
            var sw = Stopwatch.StartNew();

            try
            {
                string base64Image = Convert.ToBase64String(imageBytes);
                string mediaType = DetectMediaType(imageBytes);

                // Build the API request body
                var requestBody = new JObject
                {
                    ["model"] = _config.ModelId,
                    ["max_tokens"] = _config.MaxTokens,
                    ["system"] = systemPrompt,
                    ["messages"] = new JArray
                    {
                        new JObject
                        {
                            ["role"] = "user",
                            ["content"] = new JArray
                            {
                                new JObject
                                {
                                    ["type"] = "image",
                                    ["source"] = new JObject
                                    {
                                        ["type"] = "base64",
                                        ["media_type"] = mediaType,
                                        ["data"] = base64Image
                                    }
                                },
                                new JObject
                                {
                                    ["type"] = "text",
                                    ["text"] = userPrompt
                                }
                            }
                        }
                    }
                };

                string jsonBody = requestBody.ToString(Formatting.None);

                // Make the HTTP request
                var request = (HttpWebRequest)WebRequest.Create(ApiUrl);
                request.Method = "POST";
                request.ContentType = "application/json";
                request.Headers.Add("x-api-key", _config.ApiKey);
                request.Headers.Add("anthropic-version", ApiVersion);
                request.Timeout = _config.TimeoutSeconds * 1000;

                byte[] bodyBytes = Encoding.UTF8.GetBytes(jsonBody);
                request.ContentLength = bodyBytes.Length;

                using (var requestStream = await Task.Factory.FromAsync(
                    request.BeginGetRequestStream, request.EndGetRequestStream, null))
                {
                    await requestStream.WriteAsync(bodyBytes, 0, bodyBytes.Length);
                }

                using (var response = (HttpWebResponse)await Task.Factory.FromAsync(
                    request.BeginGetResponse, request.EndGetResponse, null))
                using (var responseStream = response.GetResponseStream())
                using (var reader = new StreamReader(responseStream, Encoding.UTF8))
                {
                    string responseJson = await reader.ReadToEndAsync();
                    sw.Stop();

                    return ParseApiResponse(responseJson, fullPage, sw.Elapsed);
                }
            }
            catch (WebException wex)
            {
                sw.Stop();
                string errorBody = "";
                if (wex.Response != null)
                {
                    using (var errorStream = wex.Response.GetResponseStream())
                    using (var reader = new StreamReader(errorStream))
                    {
                        errorBody = reader.ReadToEnd();
                    }
                }

                return new VisionAnalysisResult
                {
                    Success = false,
                    ErrorMessage = $"Claude API error: {wex.Message}. {errorBody}",
                    Duration = sw.Elapsed
                };
            }
            catch (Exception ex)
            {
                sw.Stop();
                return new VisionAnalysisResult
                {
                    Success = false,
                    ErrorMessage = $"Claude Vision service error: {ex.Message}",
                    Duration = sw.Elapsed
                };
            }
        }

        private VisionAnalysisResult ParseApiResponse(string responseJson, bool fullPage, TimeSpan duration)
        {
            try
            {
                var responseObj = JObject.Parse(responseJson);

                // Extract usage for cost tracking
                int inputTokens = responseObj["usage"]?.Value<int>("input_tokens") ?? 0;
                int outputTokens = responseObj["usage"]?.Value<int>("output_tokens") ?? 0;

                // Cost calculation (Claude Sonnet 4.5 pricing: $3/MTok input, $15/MTok output)
                decimal cost = (inputTokens * 3.0m / 1_000_000m) + (outputTokens * 15.0m / 1_000_000m);
                _sessionCost += cost;

                // Extract the text content from the response
                var content = responseObj["content"] as JArray;
                if (content == null || content.Count == 0)
                {
                    return new VisionAnalysisResult
                    {
                        Success = false,
                        ErrorMessage = "Empty response from Claude API",
                        RawJson = responseJson,
                        InputTokens = inputTokens,
                        OutputTokens = outputTokens,
                        CostUsd = cost,
                        Duration = duration
                    };
                }

                string textContent = content[0].Value<string>("text") ?? "";

                // Parse the structured response
                VisionAnalysisResult result = fullPage
                    ? _parser.ParseFullPageResponse(textContent)
                    : _parser.ParseTitleBlockResponse(textContent);

                result.InputTokens = inputTokens;
                result.OutputTokens = outputTokens;
                result.CostUsd = cost;
                result.Duration = duration;

                return result;
            }
            catch (Exception ex)
            {
                return new VisionAnalysisResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to parse Claude API response: {ex.Message}",
                    RawJson = responseJson,
                    Duration = duration
                };
            }
        }

        /// <summary>
        /// Detect image media type from magic bytes.
        /// </summary>
        private static string DetectMediaType(byte[] imageBytes)
        {
            if (imageBytes.Length >= 8 &&
                imageBytes[0] == 0x89 && imageBytes[1] == 0x50 &&
                imageBytes[2] == 0x4E && imageBytes[3] == 0x47)
                return "image/png";

            if (imageBytes.Length >= 2 &&
                imageBytes[0] == 0xFF && imageBytes[1] == 0xD8)
                return "image/jpeg";

            // Default to PNG
            return "image/png";
        }
    }
}
