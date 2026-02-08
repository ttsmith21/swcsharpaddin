namespace NM.Core.Pdf.Models
{
    /// <summary>
    /// A routing suggestion derived from a drawing note.
    /// Maps a manufacturing note to a specific work center / operation.
    /// </summary>
    public sealed class RoutingHint
    {
        /// <summary>The operation category (Deburr, Finish, Weld, etc.).</summary>
        public RoutingOp Operation { get; set; }

        /// <summary>Work center code (e.g., "F210", "F400"). Null if outside process.</summary>
        public string WorkCenter { get; set; }

        /// <summary>Routing note text for the ERP (e.g., "BREAK ALL EDGES").</summary>
        public string NoteText { get; set; }

        /// <summary>The original drawing note that generated this hint.</summary>
        public string SourceNote { get; set; }

        /// <summary>Confidence that this mapping is correct (0.0 - 1.0).</summary>
        public double Confidence { get; set; }
    }

    public enum RoutingOp
    {
        Deburr,
        Finish,
        HeatTreat,
        Weld,
        Tap,
        Drill,
        Machine,
        Inspect,
        Hardware,
        ProcessOverride,
        OutsideProcess
    }
}
