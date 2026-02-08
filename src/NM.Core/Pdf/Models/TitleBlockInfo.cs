using System;

namespace NM.Core.Pdf.Models
{
    /// <summary>
    /// Parsed title block data from an engineering drawing.
    /// Intermediate structure used by TitleBlockParser before populating DrawingData.
    /// </summary>
    public sealed class TitleBlockInfo
    {
        public string PartNumber { get; set; }
        public string Description { get; set; }
        public string Revision { get; set; }
        public string Material { get; set; }
        public string Finish { get; set; }
        public string DrawnBy { get; set; }
        public string CheckedBy { get; set; }
        public string ApprovedBy { get; set; }
        public DateTime? Date { get; set; }
        public string Scale { get; set; }
        public string Sheet { get; set; }
        public string ToleranceGeneral { get; set; }
        public string CompanyName { get; set; }

        /// <summary>Per-field confidence scores (0.0 - 1.0).</summary>
        public double PartNumberConfidence { get; set; }
        public double MaterialConfidence { get; set; }
        public double RevisionConfidence { get; set; }
        public double DescriptionConfidence { get; set; }

        public double OverallConfidence
        {
            get
            {
                int count = 0;
                double sum = 0;
                if (!string.IsNullOrEmpty(PartNumber)) { sum += PartNumberConfidence; count++; }
                if (!string.IsNullOrEmpty(Material)) { sum += MaterialConfidence; count++; }
                if (!string.IsNullOrEmpty(Revision)) { sum += RevisionConfidence; count++; }
                if (!string.IsNullOrEmpty(Description)) { sum += DescriptionConfidence; count++; }
                return count > 0 ? sum / count : 0;
            }
        }
    }
}
