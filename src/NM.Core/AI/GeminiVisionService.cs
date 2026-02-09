using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NM.Core.AI.Models;

namespace NM.Core.AI
{
    /// <summary>
    /// AI drawing analysis using the Google Gemini API with vision capabilities.
    /// Supports structured output (response_mime_type: application/json) and logprobs
    /// for calibrated per-field confidence scoring.
    /// Uses raw HttpWebRequest for .NET Framework 4.8 compatibility.
    /// </summary>
    public sealed class GeminiVisionService : IDrawingVisionService
    {
        private const string ApiUrlTemplate = "https://generativelanguage.googleapis.com/v1beta/models/{0}:generateContent?key={1}";

        private readonly VisionConfig _config;
        private readonly VisionResponseParser _parser;
        private decimal _sessionCost;

        public GeminiVisionService(VisionConfig config)
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
                    case AnalysisTier.TitleBlockOnly: return 0.002m;
                    case AnalysisTier.FullPage: return 0.010m;
                    default: return 0m;
                }
            }
        }

        public decimal SessionCost => _sessionCost;

        /// <summary>
        /// The model ID being used (e.g., "gemini-2.5-flash", "gemini-2.5-pro").
        /// </summary>
        public string ModelId => _config.ModelId;

        public async Task<VisionAnalysisResult> AnalyzeDrawingPageAsync(byte[] imageBytes, VisionContext context)
        {
            if (!IsAvailable)
            {
                return new VisionAnalysisResult
                {
                    Success = false,
                    ErrorMessage = _sessionCost >= _config.MaxSessionCost
                        ? $"Session cost limit reached (${_sessionCost:F3} >= ${_config.MaxSessionCost:F2})"
                        : "Gemini Vision service not configured (missing API key)"
                };
            }

            string prompt = (context != null && context.HasContext)
                ? VisionPrompts.GetFullPagePromptWithContext(context)
                : VisionPrompts.GetFullPagePrompt();

            return await CallGeminiApiAsync(imageBytes, prompt, fullPage: true);
        }

        public async Task<VisionAnalysisResult> AnalyzeTitleBlockAsync(byte[] imageBytes)
        {
            if (!IsAvailable)
            {
                return new VisionAnalysisResult
                {
                    Success = false,
                    ErrorMessage = "Gemini Vision service not available"
                };
            }

            string prompt = VisionPrompts.GetTitleBlockPrompt();
            return await CallGeminiApiAsync(imageBytes, prompt, fullPage: false);
        }

        /// <summary>
        /// Analyzes an image with a specific prompt (used by multi-pass extraction).
        /// Returns the raw VisionAnalysisResult with logprob-based confidence.
        /// </summary>
        public async Task<VisionAnalysisResult> AnalyzeWithPromptAsync(byte[] imageBytes, string prompt, bool fullPage)
        {
            if (!IsAvailable)
            {
                return new VisionAnalysisResult
                {
                    Success = false,
                    ErrorMessage = "Gemini Vision service not available"
                };
            }

            return await CallGeminiApiAsync(imageBytes, prompt, fullPage);
        }

        /// <summary>
        /// True if the service accepts PDF natively (no PNG rendering needed).
        /// Gemini supports application/pdf inline data.
        /// </summary>
        public bool AcceptsPdfNatively => true;

        private async Task<VisionAnalysisResult> CallGeminiApiAsync(byte[] imageBytes, string userPrompt, bool fullPage)
        {
            var sw = Stopwatch.StartNew();

            try
            {
                string base64Image = Convert.ToBase64String(imageBytes);
                string mediaType = DetectMediaType(imageBytes);

                string apiUrl = string.Format(ApiUrlTemplate, _config.ModelId, _config.ApiKey);

                // Build Gemini API request
                var requestBody = new JObject
                {
                    ["contents"] = new JArray
                    {
                        new JObject
                        {
                            ["parts"] = new JArray
                            {
                                new JObject
                                {
                                    ["inline_data"] = new JObject
                                    {
                                        ["mime_type"] = mediaType,
                                        ["data"] = base64Image
                                    }
                                },
                                new JObject
                                {
                                    ["text"] = VisionPrompts.GetSystemPrompt() + "\n\n" + userPrompt
                                }
                            }
                        }
                    },
                    ["generationConfig"] = new JObject
                    {
                        ["responseMimeType"] = "application/json",
                        ["maxOutputTokens"] = _config.MaxTokens,
                        ["temperature"] = 0.1,
                        ["responseLogprobs"] = true
                    }
                };

                string jsonBody = requestBody.ToString(Formatting.None);

                var request = (HttpWebRequest)WebRequest.Create(apiUrl);
                request.Method = "POST";
                request.ContentType = "application/json";
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
                    return ParseGeminiResponse(responseJson, fullPage, sw.Elapsed);
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
                    ErrorMessage = $"Gemini API error: {wex.Message}. {errorBody}",
                    Duration = sw.Elapsed
                };
            }
            catch (Exception ex)
            {
                sw.Stop();
                return new VisionAnalysisResult
                {
                    Success = false,
                    ErrorMessage = $"Gemini Vision service error: {ex.Message}",
                    Duration = sw.Elapsed
                };
            }
        }

        private VisionAnalysisResult ParseGeminiResponse(string responseJson, bool fullPage, TimeSpan duration)
        {
            try
            {
                var responseObj = JObject.Parse(responseJson);

                // Extract usage metadata
                var usageMetadata = responseObj["usageMetadata"];
                int inputTokens = usageMetadata?.Value<int>("promptTokenCount") ?? 0;
                int outputTokens = usageMetadata?.Value<int>("candidatesTokenCount") ?? 0;

                // Cost calculation (Gemini pricing varies by model)
                decimal cost = CalculateCost(inputTokens, outputTokens);
                _sessionCost += cost;

                // Extract text content from Gemini's response structure
                var candidates = responseObj["candidates"] as JArray;
                if (candidates == null || candidates.Count == 0)
                {
                    return new VisionAnalysisResult
                    {
                        Success = false,
                        ErrorMessage = "Empty response from Gemini API",
                        RawJson = responseJson,
                        InputTokens = inputTokens,
                        OutputTokens = outputTokens,
                        CostUsd = cost,
                        Duration = duration
                    };
                }

                var content = candidates[0]["content"];
                var parts = content?["parts"] as JArray;
                string textContent = parts?[0]?.Value<string>("text") ?? "";

                // Extract logprobs for confidence scoring
                var logprobsResult = candidates[0]["logprobsResult"];
                var perFieldConfidence = ExtractLogprobConfidence(logprobsResult, textContent);

                // Parse the structured response
                VisionAnalysisResult result = fullPage
                    ? _parser.ParseFullPageResponse(textContent)
                    : _parser.ParseTitleBlockResponse(textContent);

                // Apply logprob-based confidence scores (overriding hardcoded defaults)
                ApplyLogprobConfidence(result, perFieldConfidence);

                result.InputTokens = inputTokens;
                result.OutputTokens = outputTokens;
                result.CostUsd = cost;
                result.Duration = duration;
                result.RawJson = responseJson;

                return result;
            }
            catch (Exception ex)
            {
                return new VisionAnalysisResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to parse Gemini API response: {ex.Message}",
                    RawJson = responseJson,
                    Duration = duration
                };
            }
        }

        /// <summary>
        /// Extracts per-field confidence from Gemini's logprobs output.
        /// Logprob of -0.01 = ~99% confident; -2.0 = ~13% confident.
        /// </summary>
        private Dictionary<string, double> ExtractLogprobConfidence(JToken logprobsResult, string textContent)
        {
            var confidence = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            if (logprobsResult == null) return confidence;

            try
            {
                var topLogprobs = logprobsResult["chosenCandidates"] as JArray;
                if (topLogprobs == null) return confidence;

                // Track which JSON field we're currently inside
                string currentField = null;
                var fieldLogprobs = new List<double>();

                foreach (var token in topLogprobs)
                {
                    string tokenStr = token.Value<string>("token") ?? "";
                    double logprob = token.Value<double>("logProb");

                    // Detect JSON field names (tokens like "part_number", "material", etc.)
                    if (tokenStr.StartsWith("\"") && tokenStr.EndsWith("\"") && tokenStr.Length > 2)
                    {
                        // Save previous field's confidence
                        if (currentField != null && fieldLogprobs.Count > 0)
                        {
                            confidence[currentField] = AverageLogprobToConfidence(fieldLogprobs);
                        }

                        string potentialField = tokenStr.Trim('"');
                        // Only track known fields
                        if (IsKnownField(potentialField))
                        {
                            currentField = potentialField;
                            fieldLogprobs = new List<double>();
                        }
                        else
                        {
                            currentField = null;
                        }
                    }
                    else if (currentField != null && tokenStr != ":" && tokenStr != "," &&
                             tokenStr != "{" && tokenStr != "}" && tokenStr != "[" && tokenStr != "]" &&
                             tokenStr != "\"" && tokenStr != "\n")
                    {
                        fieldLogprobs.Add(logprob);
                    }
                }

                // Save last field
                if (currentField != null && fieldLogprobs.Count > 0)
                {
                    confidence[currentField] = AverageLogprobToConfidence(fieldLogprobs);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Gemini] Logprob extraction failed: {ex.Message}");
            }

            return confidence;
        }

        private static bool IsKnownField(string field)
        {
            switch (field)
            {
                case "part_number":
                case "description":
                case "revision":
                case "material":
                case "finish":
                case "drawn_by":
                case "date":
                case "scale":
                case "sheet":
                case "tolerance_general":
                case "text":
                case "category":
                case "type":
                case "tolerance":
                case "overall_length_inches":
                case "overall_width_inches":
                case "thickness_inches":
                    return true;
                default:
                    return false;
            }
        }

        private static double AverageLogprobToConfidence(List<double> logprobs)
        {
            if (logprobs.Count == 0) return 0.5;
            double avgLogprob = logprobs.Average();
            // Convert logprob to probability: exp(logprob)
            double confidence = Math.Exp(avgLogprob);
            return Math.Max(0.0, Math.Min(1.0, confidence));
        }

        private void ApplyLogprobConfidence(VisionAnalysisResult result, Dictionary<string, double> confidence)
        {
            if (confidence.Count == 0) return;

            ApplyFieldConfidence(result.PartNumber, "part_number", confidence);
            ApplyFieldConfidence(result.Description, "description", confidence);
            ApplyFieldConfidence(result.Revision, "revision", confidence);
            ApplyFieldConfidence(result.Material, "material", confidence);
            ApplyFieldConfidence(result.Finish, "finish", confidence);
            ApplyFieldConfidence(result.DrawnBy, "drawn_by", confidence);
            ApplyFieldConfidence(result.Date, "date", confidence);
            ApplyFieldConfidence(result.Scale, "scale", confidence);
            ApplyFieldConfidence(result.Sheet, "sheet", confidence);
            ApplyFieldConfidence(result.ToleranceGeneral, "tolerance_general", confidence);
        }

        private static void ApplyFieldConfidence(FieldResult field, string key, Dictionary<string, double> confidence)
        {
            if (field == null || !field.HasValue) return;
            if (confidence.TryGetValue(key, out double conf))
            {
                field.Confidence = conf;
            }
        }

        private decimal CalculateCost(int inputTokens, int outputTokens)
        {
            // Pricing varies by model; use conservative estimates
            string model = _config.ModelId ?? "";
            if (model.Contains("flash"))
            {
                // Gemini Flash: $0.075/MTok input, $0.30/MTok output (approx)
                return (inputTokens * 0.075m / 1_000_000m) + (outputTokens * 0.30m / 1_000_000m);
            }
            else if (model.Contains("pro"))
            {
                // Gemini Pro: $1.25/MTok input, $5.00/MTok output (approx)
                return (inputTokens * 1.25m / 1_000_000m) + (outputTokens * 5.0m / 1_000_000m);
            }
            else
            {
                // Default to Flash pricing
                return (inputTokens * 0.075m / 1_000_000m) + (outputTokens * 0.30m / 1_000_000m);
            }
        }

        private static string DetectMediaType(byte[] imageBytes)
        {
            if (imageBytes.Length >= 8 &&
                imageBytes[0] == 0x89 && imageBytes[1] == 0x50 &&
                imageBytes[2] == 0x4E && imageBytes[3] == 0x47)
                return "image/png";

            if (imageBytes.Length >= 2 &&
                imageBytes[0] == 0xFF && imageBytes[1] == 0xD8)
                return "image/jpeg";

            // PDF magic bytes: %PDF
            if (imageBytes.Length >= 4 &&
                imageBytes[0] == 0x25 && imageBytes[1] == 0x50 &&
                imageBytes[2] == 0x44 && imageBytes[3] == 0x46)
                return "application/pdf";

            return "image/png";
        }
    }
}
