namespace NM.Core.Pdf.Models
{
    /// <summary>
    /// A single row from a BOM table extracted from a PDF drawing.
    /// </summary>
    public sealed class BomEntry
    {
        public string ItemNumber { get; set; }
        public string PartNumber { get; set; }
        public string Description { get; set; }
        public int Quantity { get; set; }
        public string Material { get; set; }
        public double Confidence { get; set; }
    }
}
