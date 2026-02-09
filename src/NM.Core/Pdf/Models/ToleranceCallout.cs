using System.Collections.Generic;

namespace NM.Core.Pdf.Models
{
    /// <summary>
    /// A specific tolerance callout that is tighter than the drawing's general tolerance.
    /// </summary>
    public sealed class ToleranceCallout
    {
        /// <summary>The dimension value (e.g., "3.500").</summary>
        public string DimensionValue { get; set; }

        /// <summary>The tolerance specification (e.g., "+/-0.002", "3.500/3.502").</summary>
        public string Tolerance { get; set; }

        /// <summary>Type of tolerance: Bilateral, Unilateral, Limit, Basic.</summary>
        public ToleranceType Type { get; set; }

        /// <summary>Description of the feature (e.g., "bore diameter", "hole spacing").</summary>
        public string FeatureDescription { get; set; }

        /// <summary>True if this tolerance is confirmed tighter than the general tolerance.</summary>
        public bool IsTighterThanGeneral { get; set; }

        /// <summary>Confidence score (0.0 - 1.0).</summary>
        public double Confidence { get; set; }

        /// <summary>Page number where the callout was found.</summary>
        public int PageNumber { get; set; }
    }

    public enum ToleranceType
    {
        Bilateral,
        Unilateral,
        Limit,
        Basic,
        Unknown
    }

    /// <summary>
    /// Parsed general tolerance from the title block "UNLESS OTHERWISE SPECIFIED" block.
    /// </summary>
    public sealed class GeneralTolerance
    {
        /// <summary>Raw text of the tolerance block.</summary>
        public string RawText { get; set; }

        /// <summary>Tolerance for 2-place decimals (e.g., .XX +-0.010).</summary>
        public double? TwoPlaceDecimal { get; set; }

        /// <summary>Tolerance for 3-place decimals (e.g., .XXX +-0.005).</summary>
        public double? ThreePlaceDecimal { get; set; }

        /// <summary>Tolerance for 4-place decimals (e.g., .XXXX +-0.0005).</summary>
        public double? FourPlaceDecimal { get; set; }

        /// <summary>Angular tolerance in degrees (e.g., +-1).</summary>
        public double? AnglesDegrees { get; set; }

        /// <summary>Fractional tolerance (e.g., +-1/64).</summary>
        public string Fractional { get; set; }

        /// <summary>True if the general tolerance was successfully parsed.</summary>
        public bool IsParsed => TwoPlaceDecimal.HasValue || ThreePlaceDecimal.HasValue;

        /// <summary>
        /// Checks if a specific tolerance value is tighter than the general tolerance
        /// for the given number of decimal places.
        /// </summary>
        public bool IsTighter(double toleranceValue, int decimalPlaces)
        {
            double? generalTol = null;
            switch (decimalPlaces)
            {
                case 2: generalTol = TwoPlaceDecimal; break;
                case 3: generalTol = ThreePlaceDecimal; break;
                case 4: generalTol = FourPlaceDecimal; break;
            }

            if (!generalTol.HasValue) return false;
            return toleranceValue < generalTol.Value;
        }
    }
}
