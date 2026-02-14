using System.Collections.Generic;

namespace NM.Core.Rename
{
    /// <summary>
    /// Result summary from a batch rename operation.
    /// </summary>
    public sealed class RenameWizardResult
    {
        public int TotalComponents { get; set; }
        public int Renamed { get; set; }
        public int Skipped { get; set; }
        public int Failed { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
        public string Summary { get; set; }
    }
}
