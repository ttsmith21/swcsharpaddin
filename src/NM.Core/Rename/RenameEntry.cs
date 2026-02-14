namespace NM.Core.Rename
{
    /// <summary>
    /// Tracks a single component rename: current name, AI prediction, user-chosen final name, and confidence.
    /// </summary>
    public sealed class RenameEntry
    {
        public int Index { get; set; }
        public string CurrentFileName { get; set; }
        public string CurrentFilePath { get; set; }
        public string PredictedName { get; set; }
        public string FinalName { get; set; }
        public double Confidence { get; set; }
        public string MatchReason { get; set; }
        public BomRow MatchedBomRow { get; set; }
        public string Configuration { get; set; }
        public bool IsApproved { get; set; }
    }
}
