using NM.Core.Reconciliation.Models;

namespace NM.Core.Writeback.Models
{
    /// <summary>
    /// Result of writing a single property to the cache.
    /// </summary>
    public sealed class WritebackEntry
    {
        /// <summary>The SolidWorks custom property name.</summary>
        public string PropertyName { get; set; }

        /// <summary>Value before the write.</summary>
        public string OldValue { get; set; }

        /// <summary>Value written (or attempted).</summary>
        public string NewValue { get; set; }

        /// <summary>Outcome of the write attempt.</summary>
        public WritebackStatus Status { get; set; }

        /// <summary>Reason for skip or failure (null if applied).</summary>
        public string Reason { get; set; }

        /// <summary>Property category for display grouping.</summary>
        public PropertyCategory Category { get; set; }

        /// <summary>True if the value was changed.</summary>
        public bool IsChanged => Status == WritebackStatus.Applied &&
            !string.Equals(OldValue ?? "", NewValue ?? "", System.StringComparison.OrdinalIgnoreCase);
    }

    public enum WritebackStatus
    {
        /// <summary>Value was successfully written to the cache.</summary>
        Applied,

        /// <summary>Write was skipped (value unchanged, empty name, etc.).</summary>
        Skipped,

        /// <summary>Write failed (exception during cache update).</summary>
        Failed
    }
}
