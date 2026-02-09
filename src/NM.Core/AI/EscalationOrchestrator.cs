using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using NM.Core.AI.Models;

namespace NM.Core.AI
{
    /// <summary>
    /// Manages confidence-based escalation between AI tiers:
    ///   Tier 1: PdfPig text extraction (free, offline)
    ///   Tier 2: Gemini Flash (fast, cheap, ~88% accuracy)
    ///   Tier 3: Gemini Pro (slower, costlier, higher accuracy on hard pages)
    ///
    /// When per-field confidence from Tier 2 is low, escalates specific fields to Tier 3.
    /// </summary>
    public sealed class EscalationOrchestrator
    {
        private readonly IDrawingVisionService _tier2Service;
        private readonly IDrawingVisionService _tier3Service;
        private readonly EscalationConfig _config;
        private readonly List<EscalationEvent> _log = new List<EscalationEvent>();

        public EscalationOrchestrator(
            IDrawingVisionService tier2Service,
            IDrawingVisionService tier3Service = null,
            EscalationConfig config = null)
        {
            _tier2Service = tier2Service ?? throw new ArgumentNullException(nameof(tier2Service));
            _tier3Service = tier3Service;
            _config = config ?? new EscalationConfig();
        }

        /// <summary>
        /// Escalation event log for analysis and tuning.
        /// </summary>
        public IReadOnlyList<EscalationEvent> EscalationLog => _log.AsReadOnly();

        /// <summary>
        /// Analyzes a drawing with confidence-based escalation.
        /// Starts at Tier 2 (Gemini Flash); escalates low-confidence fields to Tier 3.
        /// </summary>
        public async Task<EscalationResult> AnalyzeWithEscalationAsync(
            byte[] fullPageImage,
            byte[] titleBlockImage,
            VisionContext context = null)
        {
            var result = new EscalationResult();

            // Tier 2: Fast analysis
            if (!_tier2Service.IsAvailable)
            {
                result.FinalTier = 0;
                result.Result = new VisionAnalysisResult
                {
                    Success = false,
                    ErrorMessage = "Tier 2 service not available"
                };
                return result;
            }

            var tier2Result = (titleBlockImage != null && titleBlockImage.Length > 0)
                ? await _tier2Service.AnalyzeTitleBlockAsync(titleBlockImage)
                : await _tier2Service.AnalyzeDrawingPageAsync(fullPageImage, context);

            result.Tier2Result = tier2Result;
            result.Tier2Cost = tier2Result?.CostUsd ?? 0;

            if (tier2Result == null || !tier2Result.Success)
            {
                result.FinalTier = 2;
                result.Result = tier2Result;
                return result;
            }

            // Check if any critical fields need escalation
            var lowConfidenceFields = FindLowConfidenceFields(tier2Result);

            if (lowConfidenceFields.Count == 0)
            {
                // All fields confident — accept Tier 2 result
                result.FinalTier = 2;
                result.Result = tier2Result;
                return result;
            }

            // Log escalation
            _log.Add(new EscalationEvent
            {
                Timestamp = DateTime.UtcNow,
                FromTier = 2,
                ToTier = 3,
                LowConfidenceFields = lowConfidenceFields.Select(f => f.Key).ToList(),
                Reason = string.Join(", ", lowConfidenceFields.Select(f =>
                    $"{f.Key}: {f.Value:F2} < {_config.GetThreshold(f.Key):F2}"))
            });

            // Tier 3: Escalation (if available)
            if (_tier3Service == null || !_tier3Service.IsAvailable)
            {
                result.FinalTier = 2;
                result.Result = tier2Result;
                result.EscalationSkipped = true;
                result.EscalationReason = "Tier 3 service not available";
                return result;
            }

            var tier3Result = await _tier3Service.AnalyzeDrawingPageAsync(fullPageImage, context);
            result.Tier3Result = tier3Result;
            result.Tier3Cost = tier3Result?.CostUsd ?? 0;

            if (tier3Result == null || !tier3Result.Success)
            {
                result.FinalTier = 2;
                result.Result = tier2Result;
                return result;
            }

            // Merge: use Tier 3 values for low-confidence fields, keep Tier 2 for the rest
            var merged = MergeResults(tier2Result, tier3Result, lowConfidenceFields);
            result.FinalTier = 3;
            result.Result = merged;

            return result;
        }

        private Dictionary<string, double> FindLowConfidenceFields(VisionAnalysisResult result)
        {
            var lowFields = new Dictionary<string, double>();

            CheckFieldConfidence(result.PartNumber, "part_number", lowFields);
            CheckFieldConfidence(result.Material, "material", lowFields);
            CheckFieldConfidence(result.Revision, "revision", lowFields);
            CheckFieldConfidence(result.Description, "description", lowFields);
            CheckFieldConfidence(result.Finish, "finish", lowFields);
            CheckFieldConfidence(result.ToleranceGeneral, "tolerance_general", lowFields);

            // Manufacturing notes with low average confidence
            if (result.ManufacturingNotes.Count > 0)
            {
                double avgNoteConfidence = result.ManufacturingNotes.Average(n => n.Confidence);
                if (avgNoteConfidence < _config.NotesConfidenceThreshold)
                {
                    lowFields["manufacturing_notes"] = avgNoteConfidence;
                }
            }

            // GD&T with low confidence
            if (result.GdtCallouts.Count > 0)
            {
                double avgGdtConfidence = result.GdtCallouts.Average(g => g.Confidence);
                if (avgGdtConfidence < _config.GdtConfidenceThreshold)
                {
                    lowFields["gdt_callouts"] = avgGdtConfidence;
                }
            }

            return lowFields;
        }

        private void CheckFieldConfidence(FieldResult field, string fieldName, Dictionary<string, double> lowFields)
        {
            if (field == null || !field.HasValue) return;
            double threshold = _config.GetThreshold(fieldName);
            if (field.Confidence < threshold)
            {
                lowFields[fieldName] = field.Confidence;
            }
        }

        private VisionAnalysisResult MergeResults(
            VisionAnalysisResult tier2,
            VisionAnalysisResult tier3,
            Dictionary<string, double> lowConfidenceFields)
        {
            // Start with Tier 2 as base
            var merged = tier2;

            // Override low-confidence fields with Tier 3 values
            if (lowConfidenceFields.ContainsKey("part_number") && tier3.PartNumber?.HasValue == true)
                merged.PartNumber = tier3.PartNumber;
            if (lowConfidenceFields.ContainsKey("material") && tier3.Material?.HasValue == true)
                merged.Material = tier3.Material;
            if (lowConfidenceFields.ContainsKey("revision") && tier3.Revision?.HasValue == true)
                merged.Revision = tier3.Revision;
            if (lowConfidenceFields.ContainsKey("description") && tier3.Description?.HasValue == true)
                merged.Description = tier3.Description;
            if (lowConfidenceFields.ContainsKey("finish") && tier3.Finish?.HasValue == true)
                merged.Finish = tier3.Finish;
            if (lowConfidenceFields.ContainsKey("tolerance_general") && tier3.ToleranceGeneral?.HasValue == true)
                merged.ToleranceGeneral = tier3.ToleranceGeneral;

            // For notes and GD&T, prefer the Tier 3 result entirely if escalated
            if (lowConfidenceFields.ContainsKey("manufacturing_notes") && tier3.ManufacturingNotes.Count > 0)
            {
                merged.ManufacturingNotes.Clear();
                merged.ManufacturingNotes.AddRange(tier3.ManufacturingNotes);
            }
            if (lowConfidenceFields.ContainsKey("gdt_callouts") && tier3.GdtCallouts.Count > 0)
            {
                merged.GdtCallouts.Clear();
                merged.GdtCallouts.AddRange(tier3.GdtCallouts);
            }

            return merged;
        }
    }

    /// <summary>
    /// Configuration for escalation thresholds.
    /// </summary>
    public sealed class EscalationConfig
    {
        /// <summary>Default confidence threshold below which a field triggers escalation.</summary>
        public double DefaultThreshold { get; set; } = 0.60;

        /// <summary>Threshold for critical fields (part number, material).</summary>
        public double CriticalFieldThreshold { get; set; } = 0.70;

        /// <summary>Threshold for manufacturing notes (average confidence).</summary>
        public double NotesConfidenceThreshold { get; set; } = 0.50;

        /// <summary>Threshold for GD&T callouts.</summary>
        public double GdtConfidenceThreshold { get; set; } = 0.50;

        public double GetThreshold(string fieldName)
        {
            switch (fieldName)
            {
                case "part_number":
                case "material":
                    return CriticalFieldThreshold;
                case "manufacturing_notes":
                    return NotesConfidenceThreshold;
                case "gdt_callouts":
                    return GdtConfidenceThreshold;
                default:
                    return DefaultThreshold;
            }
        }
    }

    /// <summary>
    /// Result of escalation-aware analysis.
    /// </summary>
    public sealed class EscalationResult
    {
        /// <summary>The final merged result.</summary>
        public VisionAnalysisResult Result { get; set; }

        /// <summary>Which tier produced the final result (2 or 3).</summary>
        public int FinalTier { get; set; }

        /// <summary>Tier 2 raw result.</summary>
        public VisionAnalysisResult Tier2Result { get; set; }

        /// <summary>Tier 3 raw result (null if no escalation).</summary>
        public VisionAnalysisResult Tier3Result { get; set; }

        /// <summary>Cost of Tier 2 analysis.</summary>
        public decimal Tier2Cost { get; set; }

        /// <summary>Cost of Tier 3 analysis (0 if no escalation).</summary>
        public decimal Tier3Cost { get; set; }

        /// <summary>Total cost.</summary>
        public decimal TotalCost => Tier2Cost + Tier3Cost;

        /// <summary>True if escalation was needed but service was unavailable.</summary>
        public bool EscalationSkipped { get; set; }

        /// <summary>Reason escalation was skipped.</summary>
        public string EscalationReason { get; set; }
    }

    /// <summary>
    /// Records an escalation decision for later analysis.
    /// </summary>
    public sealed class EscalationEvent
    {
        public DateTime Timestamp { get; set; }
        public int FromTier { get; set; }
        public int ToTier { get; set; }
        public List<string> LowConfidenceFields { get; set; } = new List<string>();
        public string Reason { get; set; }
    }
}
