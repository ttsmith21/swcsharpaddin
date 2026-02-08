using NM.Core.Pdf.Models;

namespace NM.Core.Reconciliation.Models
{
    /// <summary>
    /// A suggested routing operation derived from drawing notes.
    /// </summary>
    public sealed class RoutingSuggestion
    {
        /// <summary>The operation type.</summary>
        public RoutingOp Operation { get; set; }

        /// <summary>Suggested operation sequence number (e.g., 30 for OP30).</summary>
        public int SuggestedOpNumber { get; set; }

        /// <summary>Work center code (e.g., "F210", "F400"). Null for outside process.</summary>
        public string WorkCenter { get; set; }

        /// <summary>Routing note text for the ERP.</summary>
        public string NoteText { get; set; }

        /// <summary>The original drawing note that generated this suggestion.</summary>
        public string SourceNote { get; set; }

        /// <summary>Confidence (0.0 - 1.0).</summary>
        public double Confidence { get; set; }

        /// <summary>Whether this is an addition, modification, or constraint.</summary>
        public SuggestionType Type { get; set; }
    }

    public enum SuggestionType
    {
        /// <summary>Add a new operation to the routing.</summary>
        AddOperation,

        /// <summary>Modify an existing operation (e.g., change work center).</summary>
        ModifyOperation,

        /// <summary>Add a routing note to an existing operation.</summary>
        AddNote
    }
}
