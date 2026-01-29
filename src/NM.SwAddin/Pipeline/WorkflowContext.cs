using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using NM.Core.Models;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.Pipeline
{
    /// <summary>
    /// Holds state between validation and processing passes.
    /// Created by WorkflowDispatcher, passed through the two-pass workflow.
    /// </summary>
    public sealed class WorkflowContext
    {
        /// <summary>
        /// How the workflow was initiated.
        /// </summary>
        public enum SourceType
        {
            None = 0,
            Part = 1,
            Assembly = 2,
            Drawing = 3,
            Folder = 4
        }

        /// <summary>
        /// What initiated the workflow.
        /// </summary>
        public SourceType Source { get; set; } = SourceType.None;

        /// <summary>
        /// Root folder path (for Folder source) or drawing file path.
        /// </summary>
        public string RootPath { get; set; }

        /// <summary>
        /// The original document that was open (Part, Assembly, or Drawing).
        /// Null if source is Folder or None.
        /// </summary>
        public IModelDoc2 SourceDocument { get; set; }

        /// <summary>
        /// All models collected for processing (before validation).
        /// </summary>
        public List<SwModelInfo> AllModels { get; } = new List<SwModelInfo>();

        /// <summary>
        /// Models that passed validation (ready for processing).
        /// </summary>
        public List<SwModelInfo> GoodModels { get; } = new List<SwModelInfo>();

        /// <summary>
        /// Models that failed validation (need user intervention).
        /// </summary>
        public List<SwModelInfo> ProblemModels { get; } = new List<SwModelInfo>();

        /// <summary>
        /// Models successfully processed in Pass 2.
        /// </summary>
        public List<SwModelInfo> ProcessedModels { get; } = new List<SwModelInfo>();

        /// <summary>
        /// Models that failed during processing (Pass 2 errors).
        /// </summary>
        public List<SwModelInfo> FailedModels { get; } = new List<SwModelInfo>();

        /// <summary>
        /// True after validation pass completes.
        /// </summary>
        public bool ValidationComplete { get; set; }

        /// <summary>
        /// True after processing pass completes.
        /// </summary>
        public bool ProcessingComplete { get; set; }

        /// <summary>
        /// True if user canceled during any step.
        /// </summary>
        public bool UserCanceled { get; set; }

        /// <summary>
        /// Total BOM quantity (sum of all component quantities for assemblies).
        /// </summary>
        public int TotalBomQuantity { get; set; }

        /// <summary>
        /// Timing for validation pass.
        /// </summary>
        public TimeSpan ValidationElapsed { get; set; }

        /// <summary>
        /// Timing for processing pass.
        /// </summary>
        public TimeSpan ProcessingElapsed { get; set; }

        /// <summary>
        /// Stopwatch for overall timing.
        /// </summary>
        private readonly Stopwatch _stopwatch = new Stopwatch();

        public void StartTiming() => _stopwatch.Start();
        public void StopTiming() => _stopwatch.Stop();
        public TimeSpan TotalElapsed => _stopwatch.Elapsed;

        /// <summary>
        /// Generates a summary string for the validation phase.
        /// </summary>
        public string GetValidationSummary()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Validation Summary");
            sb.AppendLine("─────────────────");
            sb.AppendLine($"Source: {Source}");
            sb.AppendLine($"Total Models: {AllModels.Count}");
            sb.AppendLine($"Valid: {GoodModels.Count}");
            sb.AppendLine($"Problems: {ProblemModels.Count}");
            sb.AppendLine($"Elapsed: {ValidationElapsed.TotalSeconds:F1}s");
            return sb.ToString();
        }

        /// <summary>
        /// Generates a summary string for the processing phase.
        /// </summary>
        public string GetProcessingSummary()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Processing Summary");
            sb.AppendLine("──────────────────");
            sb.AppendLine($"Processed OK: {ProcessedModels.Count}");
            sb.AppendLine($"Failed: {FailedModels.Count}");
            sb.AppendLine($"Elapsed: {ProcessingElapsed.TotalSeconds:F1}s");
            return sb.ToString();
        }

        /// <summary>
        /// Generates a complete summary for both phases.
        /// </summary>
        public string GetFinalSummary()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Workflow Complete");
            sb.AppendLine("═════════════════");
            sb.AppendLine($"Source: {Source}");
            if (!string.IsNullOrEmpty(RootPath))
                sb.AppendLine($"Path: {RootPath}");
            sb.AppendLine();
            sb.AppendLine($"Total Discovered: {AllModels.Count}");
            if (TotalBomQuantity > 0)
                sb.AppendLine($"Total BOM Quantity: {TotalBomQuantity}");
            sb.AppendLine($"Validated OK: {GoodModels.Count}");
            sb.AppendLine($"Validation Problems: {ProblemModels.Count}");
            sb.AppendLine($"Processed OK: {ProcessedModels.Count}");
            sb.AppendLine($"Processing Failed: {FailedModels.Count}");
            sb.AppendLine();
            sb.AppendLine($"Total Time: {TotalElapsed.TotalSeconds:F1}s");
            return sb.ToString();
        }
    }
}
