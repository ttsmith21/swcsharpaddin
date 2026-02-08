namespace NM.Core.Pdf.Models
{
    /// <summary>
    /// A note extracted from an engineering drawing.
    /// </summary>
    public sealed class DrawingNote
    {
        public string Text { get; set; }
        public NoteCategory Category { get; set; }
        public RoutingImpact Impact { get; set; }
        public double Confidence { get; set; }

        /// <summary>Page number where the note was found (1-based).</summary>
        public int PageNumber { get; set; }

        public DrawingNote() { }

        public DrawingNote(string text, NoteCategory category, double confidence = 1.0)
        {
            Text = text;
            Category = category;
            Impact = ClassifyImpact(category);
            Confidence = confidence;
        }

        private static RoutingImpact ClassifyImpact(NoteCategory category)
        {
            switch (category)
            {
                case NoteCategory.Deburr:
                case NoteCategory.Finish:
                case NoteCategory.HeatTreat:
                case NoteCategory.Weld:
                case NoteCategory.Machine:
                case NoteCategory.Hardware:
                    return RoutingImpact.AddOperation;
                case NoteCategory.ProcessConstraint:
                    return RoutingImpact.ModifyOperation;
                case NoteCategory.Inspect:
                    return RoutingImpact.AddOperation;
                default:
                    return RoutingImpact.Informational;
            }
        }
    }

    public enum NoteCategory
    {
        General,
        Deburr,
        Finish,
        HeatTreat,
        Weld,
        Machine,
        Inspect,
        Hardware,
        ProcessConstraint,
        Material,
        Dimension,
        Tolerance
    }

    public enum RoutingImpact
    {
        Informational,
        AddOperation,
        ModifyOperation,
        Constraint
    }
}
