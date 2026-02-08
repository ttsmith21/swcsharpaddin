using System.Collections.Generic;

namespace NM.Core.Reconciliation.Models
{
    /// <summary>
    /// A suggested file rename based on drawing part number.
    /// </summary>
    public sealed class RenameSuggestion
    {
        /// <summary>Current file path.</summary>
        public string OldPath { get; set; }

        /// <summary>Suggested new file path.</summary>
        public string NewPath { get; set; }

        /// <summary>Current drawing file path (if exists).</summary>
        public string OldDrawingPath { get; set; }

        /// <summary>Suggested new drawing file path.</summary>
        public string NewDrawingPath { get; set; }

        /// <summary>Reason for the rename suggestion.</summary>
        public string Reason { get; set; }

        /// <summary>Confidence (0.0 - 1.0).</summary>
        public double Confidence { get; set; }

        /// <summary>Assembly files that reference this part (would need updating).</summary>
        public List<string> AffectedAssemblies { get; } = new List<string>();

        /// <summary>Always requires explicit user approval.</summary>
        public bool RequiresUserApproval => true;
    }
}
