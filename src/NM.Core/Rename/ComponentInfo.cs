namespace NM.Core.Rename
{
    /// <summary>
    /// Lightweight info about an assembly component for matching against BOM rows.
    /// Decoupled from SolidWorks COM types so matching logic stays in NM.Core.
    /// </summary>
    public sealed class ComponentInfo
    {
        public int Index { get; set; }
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string Configuration { get; set; }
        public string Material { get; set; }
        public int InstanceCount { get; set; }
    }
}
