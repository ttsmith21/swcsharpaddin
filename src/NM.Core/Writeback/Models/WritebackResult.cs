using System.Collections.Generic;
using System.Linq;

namespace NM.Core.Writeback.Models
{
    /// <summary>
    /// Aggregate result of applying property suggestions to a model's cache.
    /// </summary>
    public sealed class WritebackResult
    {
        /// <summary>Properties that were successfully written.</summary>
        public List<WritebackEntry> Applied { get; } = new List<WritebackEntry>();

        /// <summary>Properties that were skipped (already correct, empty, etc.).</summary>
        public List<WritebackEntry> Skipped { get; } = new List<WritebackEntry>();

        /// <summary>Properties that failed to write.</summary>
        public List<WritebackEntry> Failed { get; } = new List<WritebackEntry>();

        /// <summary>True if all suggestions were applied or skipped (no failures).</summary>
        public bool Success => Failed.Count == 0;

        /// <summary>Total number of suggestions processed.</summary>
        public int TotalProcessed => Applied.Count + Skipped.Count + Failed.Count;

        /// <summary>Number of properties actually changed.</summary>
        public int ChangedCount => Applied.Count(e => e.IsChanged);

        /// <summary>Human-readable summary for display.</summary>
        public string Summary =>
            $"{Applied.Count} applied, {Skipped.Count} skipped, {Failed.Count} failed";
    }
}
