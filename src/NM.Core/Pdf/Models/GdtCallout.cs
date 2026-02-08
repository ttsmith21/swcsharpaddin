using System.Collections.Generic;

namespace NM.Core.Pdf.Models
{
    /// <summary>
    /// A GD&amp;T feature control frame extracted from a drawing.
    /// </summary>
    public sealed class GdtCallout
    {
        public string Type { get; set; }  // flatness, parallelism, position, etc.
        public string Tolerance { get; set; }
        public List<string> DatumReferences { get; } = new List<string>();
        public string FeatureDescription { get; set; }
        public double Confidence { get; set; }
    }
}
