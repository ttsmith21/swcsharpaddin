using System;
using System.IO;
using NM.Core; // for CustomPropertyData

namespace NM.Core.Models
{
    /// <summary>
    /// Lightweight model wrapper for pipeline context.
    /// </summary>
    public sealed class SwModelInfo
    {
        public enum ModelType { Unknown = 0, Part = 1, Assembly = 2, Drawing = 3 }
        public enum ProcessingState { Unprocessed = 0, Validated = 1, Processing = 2, Processed = 3, Problem = 4 }

        public string FilePath { get; }
        public string FileName => string.IsNullOrWhiteSpace(FilePath) ? string.Empty : Path.GetFileName(FilePath);

        private string _configuration = string.Empty;
        /// <summary>
        /// Active configuration. Keep CustomProperties.ConfigurationName in sync.
        /// </summary>
        public string Configuration
        {
            get => _configuration;
            set
            {
                _configuration = value ?? string.Empty;
                if (CustomProperties != null)
                {
                    CustomProperties.ConfigurationName = _configuration;
                }
            }
        }

        public ModelType Type { get; private set; }
        public ProcessingState State { get; private set; }
        public string ProblemDescription { get; private set; }

        // Unified dirty flag: model edits OR pending custom property changes.
        private bool _modelDirty;
        /// <summary>
        /// True when there are pending changes that require a save: either SOLIDWORKS model edits
        /// (tracked via MarkModelDirty/MarkModelClean) or custom property changes in the cache.
        /// </summary>
        public bool IsDirty => _modelDirty || (CustomProperties?.IsDirty ?? false);

        public DateTime? ValidatedAt { get; private set; }
        public DateTime? ProcessedAt { get; private set; }

        // cache for COM doc (object type to keep NM.Core free of COM types)
        public object ModelDoc { get; set; }

        /// <summary>
        /// Custom property cache with change tracking. Pure data; written by add-in service later.
        /// </summary>
        public CustomPropertyData CustomProperties { get; }

        /// <summary>
        /// Concise info for logs/UI: File, Config, Type, State, Dirty.
        /// </summary>
        public string Context
        {
            get
            {
                var cfg = string.IsNullOrWhiteSpace(Configuration) ? string.Empty : $" @\"{Configuration}\"";
                var dirty = IsDirty ? " | dirty" : string.Empty;
                return $"{FileName}{cfg} | {Type} | {State}{dirty}";
            }
        }

        public SwModelInfo(string filePath, string configuration = null)
        {
            if (filePath == null) filePath = string.Empty;
            FilePath = filePath;
            CustomProperties = new CustomPropertyData();
            Configuration = configuration ?? string.Empty; // syncs CustomProperties.ConfigurationName
            State = ProcessingState.Unprocessed;

            var ext = (Path.GetExtension(filePath) ?? string.Empty).ToLowerInvariant();
            if (ext == ".sldprt") Type = ModelType.Part;
            else if (ext == ".sldasm") Type = ModelType.Assembly;
            else if (ext == ".slddrw") Type = ModelType.Drawing;
            else Type = ModelType.Unknown;
        }

        public void MarkValidated(bool success, string problem = null)
        {
            if (State != ProcessingState.Unprocessed) throw new InvalidOperationException($"Cannot validate from {State}");
            ValidatedAt = DateTime.UtcNow;
            if (success) { State = ProcessingState.Validated; ProblemDescription = null; }
            else { State = ProcessingState.Problem; ProblemDescription = problem ?? "Validation failed"; }
        }

        public void StartProcessing()
        {
            if (State != ProcessingState.Validated) throw new InvalidOperationException($"Cannot start from {State}");
            State = ProcessingState.Processing;
        }

        public void CompleteProcessing(bool success = true, string problem = null)
        {
            if (State != ProcessingState.Processing) throw new InvalidOperationException($"Cannot complete from {State}");
            ProcessedAt = DateTime.UtcNow;
            if (success) { State = ProcessingState.Processed; ProblemDescription = null; }
            else { State = ProcessingState.Problem; ProblemDescription = problem ?? "Processing failed"; }
        }

        /// <summary>
        /// Mark that the underlying SOLIDWORKS document has unsaved changes (non-property edits).
        /// </summary>
        public void MarkModelDirty() => _modelDirty = true;

        /// <summary>
        /// Mark only the model-edit dirty flag as clean. Does not touch custom property cache.
        /// Call this after a successful Save when no property writes were pending or after property writes are flushed.
        /// </summary>
        public void MarkModelClean() => _modelDirty = false;

        /// <summary>
        /// Mark both model and custom property caches as clean.
        /// Use after successfully writing properties and saving the document.
        /// </summary>
        public void MarkAllClean()
        {
            _modelDirty = false;
            CustomProperties?.MarkClean();
        }

        public override string ToString() => $"{FileName} [{State}]{(IsDirty ? "*" : string.Empty)}";
    }
}
