using System;
using System.Globalization;

namespace NM.Core
{
    /// <summary>
    /// Represents a single problem part occurrence captured during processing.
    /// Pure data class with no SolidWorks dependencies.
    /// </summary>
    public sealed class ProblemPartInfo : IEquatable<ProblemPartInfo>
    {
        /// <summary>Full path to the part/assembly/drawing.</summary>
        public string PartPath { get; }
        /// <summary>Human-readable reason for the problem.</summary>
        public string Reason { get; private set; }
        /// <summary>True if the problem still needs a fix.</summary>
        public bool NeedsFix { get; private set; }
        /// <summary>Configuration name related to the problem (optional; may be empty).</summary>
        public string Configuration { get; }

        /// <summary>
        /// Create a new problem part record.
        /// </summary>
        public ProblemPartInfo(string partPath, string reason, string configuration = "")
        {
            // Use SolidWorksApiWrapper.ValidateString pattern indirectly: keep NM.Core free of COM
            if (string.IsNullOrWhiteSpace(partPath)) throw new ArgumentException("Invalid part path", nameof(partPath));
            if (string.IsNullOrWhiteSpace(reason)) throw new ArgumentException("Invalid reason", nameof(reason));

            PartPath = partPath.Trim();
            Reason = reason.Trim();
            Configuration = (configuration ?? string.Empty).Trim();
            NeedsFix = true;
        }

        /// <summary>Marks this problem as fixed and clears the reason.</summary>
        public void MarkAsFixed()
        {
            NeedsFix = false;
            Reason = string.Empty;
        }

        /// <summary>Display string of the form: "path [config] | reason".</summary>
        public string GetDisplayText()
        {
            if (!string.IsNullOrEmpty(Configuration))
            {
                return string.Concat(PartPath, " [", Configuration, "] | ", Reason);
            }
            return string.Concat(PartPath, " | ", Reason);
        }

        public override string ToString() => GetDisplayText();

        public override bool Equals(object obj) => Equals(obj as ProblemPartInfo);

        public bool Equals(ProblemPartInfo other)
        {
            if (ReferenceEquals(this, other)) return true;
            if (other == null) return false;
            return string.Equals(PartPath ?? string.Empty, other.PartPath ?? string.Empty, StringComparison.OrdinalIgnoreCase)
                && string.Equals(Configuration ?? string.Empty, other.Configuration ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var pathHash = (PartPath ?? string.Empty).ToLowerInvariant().GetHashCode();
                var cfgHash = (Configuration ?? string.Empty).ToLowerInvariant().GetHashCode();
                return (pathHash * 397) ^ cfgHash;
            }
        }
    }
}
