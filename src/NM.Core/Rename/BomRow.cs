namespace NM.Core.Rename
{
    /// <summary>
    /// A single row extracted from a BOM table on a PDF drawing.
    /// </summary>
    public sealed class BomRow
    {
        public int ItemNumber { get; set; }
        public string PartNumber { get; set; }
        public string Description { get; set; }
        public string Material { get; set; }
        public int Quantity { get; set; }
    }
}
