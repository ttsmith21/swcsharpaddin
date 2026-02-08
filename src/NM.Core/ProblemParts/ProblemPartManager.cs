using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NM.Core.Models;

namespace NM.Core.ProblemParts
{
    public sealed class ProblemPartManager
    {
        // Simple singleton for app-wide access
        private static readonly Lazy<ProblemPartManager> _instance = new Lazy<ProblemPartManager>(() => new ProblemPartManager());
        public static ProblemPartManager Instance => _instance.Value;

        private readonly List<ProblemItem> _items = new List<ProblemItem>();

        public enum ProblemCategory
        {
            SheetMetalConversion,
            GeometryValidation,
            MaterialMissing,
            ThicknessExtraction,
            FileAccess,
            Suppressed,
            Lightweight,
            Imported,
            MixedBody,
            MultiBody,
            ProcessingError,
            ClassificationFailed,
            Fatal
        }

        /// <summary>
        /// User-assigned part type override from problem parts review.
        /// When set to anything other than None, implies rbPartType=1.
        /// Maps to rbPartTypeSub: Machined=0, Purchased=1, CustomerSupplied=2.
        /// </summary>
        public enum PartTypeOverride
        {
            None = -1,
            Machined = 0,
            Purchased = 1,
            CustomerSupplied = 2
        }

        public sealed class ProblemItem
        {
            public string FilePath { get; set; }
            public string Configuration { get; set; }
            public string ComponentName { get; set; }
            public string ProblemDescription { get; set; }
            public ProblemCategory Category { get; set; }
            public int RetryCount { get; set; }
            public bool UserReviewed { get; set; }
            public DateTime FirstEncountered { get; set; }
            public DateTime LastAttempted { get; set; }
            public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            /// <summary>
            /// User-assigned part type override. When not None, implies rbPartType=1.
            /// </summary>
            public PartTypeOverride TypeOverride { get; set; } = PartTypeOverride.None;

            public string DisplayName => Path.GetFileName(FilePath ?? string.Empty);
            public bool HasTypeOverride => TypeOverride != PartTypeOverride.None;
            public bool CanRetry => Category != ProblemCategory.Fatal && RetryCount < 3;
            public string Status => UserReviewed ? "Ready to Retry" : "Pending Review";
        }

        public void AddProblemPart(string filePath, string configuration, string componentName, string reason, ProblemCategory category)
        {
            if (string.IsNullOrWhiteSpace(filePath)) return;
            var existing = _items.FirstOrDefault(p =>
                string.Equals(p.FilePath ?? string.Empty, filePath ?? string.Empty, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(p.Configuration ?? string.Empty, configuration ?? string.Empty, StringComparison.OrdinalIgnoreCase));

            if (existing != null)
            {
                existing.RetryCount = Math.Max(existing.RetryCount + 1, 1);
                existing.LastAttempted = DateTime.Now;
                // Aggregate reasons instead of replacing
                if (!string.IsNullOrEmpty(reason) && (existing.ProblemDescription == null || !existing.ProblemDescription.Contains(reason)))
                    existing.ProblemDescription = string.IsNullOrEmpty(existing.ProblemDescription)
                        ? reason
                        : existing.ProblemDescription + "; " + reason;
                existing.Category = category;
                ErrorHandler.HandleError(nameof(ProblemPartManager), $"Problem part retry #{existing.RetryCount}: {existing.DisplayName}", null, ErrorHandler.LogLevel.Warning);
            }
            else
            {
                var item = new ProblemItem
                {
                    FilePath = filePath,
                    Configuration = configuration,
                    ComponentName = componentName,
                    ProblemDescription = reason ?? string.Empty,
                    Category = category,
                    RetryCount = 1,
                    FirstEncountered = DateTime.Now,
                    LastAttempted = DateTime.Now
                };
                _items.Add(item);
                ErrorHandler.HandleError(nameof(ProblemPartManager), $"New problem part: {item.DisplayName} - {item.ProblemDescription}", null, ErrorHandler.LogLevel.Warning);
            }
        }

        public void AddProblemPart(ModelInfo modelInfo, string reason, ProblemCategory category)
        {
            if (modelInfo == null) return;
            AddProblemPart(modelInfo.FilePath, modelInfo.ConfigurationName, string.Empty, reason, category);
        }

        public System.Collections.Generic.List<ProblemItem> GetProblemParts(bool includeReviewed = true)
        {
            if (includeReviewed) return _items.ToList();
            return _items.Where(p => !p.UserReviewed).ToList();
        }

        public System.Collections.Generic.List<ProblemItem> GetRetryableParts()
            => _items.Where(p => p.CanRetry && p.UserReviewed).ToList();

        public void MarkAsReviewed(IEnumerable<ProblemItem> parts)
        {
            if (parts == null) return;
            foreach (var p in parts)
            {
                p.UserReviewed = true;
                ErrorHandler.DebugLog($"[Problems] Marked reviewed: {p.DisplayName}");
            }
        }

        public void SetTypeOverride(ProblemItem item, PartTypeOverride typeOverride)
        {
            if (item == null) return;
            item.TypeOverride = typeOverride;
            item.UserReviewed = true;
            ErrorHandler.DebugLog($"[Problems] Type override: {item.DisplayName} -> {typeOverride}");
        }

        public void RemoveResolvedPart(ProblemItem part)
        {
            if (part == null) return;
            _items.Remove(part);
            ErrorHandler.DebugLog($"[Problems] Resolved: {part.DisplayName}");
        }

        public string GenerateSummary()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Problem Parts Summary:");
            sb.AppendLine($"Total: {_items.Count}");
            var byCat = _items.GroupBy(p => p.Category).OrderByDescending(g => g.Count());
            foreach (var g in byCat)
            {
                sb.AppendLine($"  {g.Key}: {g.Count()}");
            }
            int multi = _items.Count(p => p.RetryCount > 1);
            if (multi > 0) sb.AppendLine($"\nMultiple Attempts: {multi}");
            return sb.ToString();
        }
    }
}
