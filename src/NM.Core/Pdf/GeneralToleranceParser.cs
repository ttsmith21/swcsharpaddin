using System.Text.RegularExpressions;
using NM.Core.Pdf.Models;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Parses the "UNLESS OTHERWISE SPECIFIED" general tolerance block
    /// commonly found in engineering drawing title blocks.
    /// </summary>
    public static class GeneralToleranceParser
    {
        // Pattern: .XX +-0.010 or .XX ± 0.01 or TWO PLACE DECIMAL ±.01
        private static readonly Regex TwoPlacePattern = new Regex(
            @"(?:\.XX|TWO\s*(?:PLACE|PL)\s*(?:DECIMAL|DEC))\s*[:\s]*[±+\-/]*\s*\.?(\d{2,3})\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Pattern: .XXX +-0.005 or THREE PLACE DECIMAL ±.005
        private static readonly Regex ThreePlacePattern = new Regex(
            @"(?:\.XXX|THREE\s*(?:PLACE|PL)\s*(?:DECIMAL|DEC))\s*[:\s]*[±+\-/]*\s*\.?(\d{3,4})\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Pattern: .XXXX +-0.0005
        private static readonly Regex FourPlacePattern = new Regex(
            @"(?:\.XXXX|FOUR\s*(?:PLACE|PL)\s*(?:DECIMAL|DEC))\s*[:\s]*[±+\-/]*\s*\.?(\d{4,5})\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Pattern: ANGLES +-1 DEG or ANGULAR ±0°30' or ANGLES ± 1
        private static readonly Regex AnglesPattern = new Regex(
            @"(?:ANGLE[S]?|ANGULAR)\s*[:\s]*[±+\-/]*\s*(\d+(?:\.\d+)?)\s*(?:°|DEG|DEGREES?)?",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Pattern: FRACTIONAL +-1/64 or FRACTIONS ± 1/32
        private static readonly Regex FractionalPattern = new Regex(
            @"(?:FRACTION(?:AL|S)?)\s*[:\s]*[±+\-/]*\s*(\d+/\d+)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Fallback: match tolerance values like ±.010 or +/-.005 near relevant keywords
        private static readonly Regex GenericTolerancePattern = new Regex(
            @"[±]\s*\.?(\d{2,5})\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Parses a general tolerance text string into structured data.
        /// </summary>
        public static GeneralTolerance Parse(string toleranceText)
        {
            var result = new GeneralTolerance { RawText = toleranceText };

            if (string.IsNullOrWhiteSpace(toleranceText))
                return result;

            // Try specific patterns first
            var m2 = TwoPlacePattern.Match(toleranceText);
            if (m2.Success)
            {
                result.TwoPlaceDecimal = ParseToleranceValue(m2.Groups[1].Value);
            }

            var m3 = ThreePlacePattern.Match(toleranceText);
            if (m3.Success)
            {
                result.ThreePlaceDecimal = ParseToleranceValue(m3.Groups[1].Value);
            }

            var m4 = FourPlacePattern.Match(toleranceText);
            if (m4.Success)
            {
                result.FourPlaceDecimal = ParseToleranceValue(m4.Groups[1].Value);
            }

            var mAngle = AnglesPattern.Match(toleranceText);
            if (mAngle.Success)
            {
                if (double.TryParse(mAngle.Groups[1].Value, out double angle))
                    result.AnglesDegrees = angle;
            }

            var mFrac = FractionalPattern.Match(toleranceText);
            if (mFrac.Success)
            {
                result.Fractional = mFrac.Groups[1].Value;
            }

            // If we didn't get specific patterns, try to extract from generic ± values
            if (!result.TwoPlaceDecimal.HasValue && !result.ThreePlaceDecimal.HasValue)
            {
                var genericMatches = GenericTolerancePattern.Matches(toleranceText);
                foreach (Match gm in genericMatches)
                {
                    double? val = ParseToleranceValue(gm.Groups[1].Value);
                    if (!val.HasValue) continue;

                    string digits = gm.Groups[1].Value.TrimStart('0');
                    if (digits.Length <= 2 && !result.TwoPlaceDecimal.HasValue)
                        result.TwoPlaceDecimal = val;
                    else if (digits.Length == 3 && !result.ThreePlaceDecimal.HasValue)
                        result.ThreePlaceDecimal = val;
                    else if (digits.Length >= 4 && !result.FourPlaceDecimal.HasValue)
                        result.FourPlaceDecimal = val;
                }
            }

            return result;
        }

        private static double? ParseToleranceValue(string digits)
        {
            if (string.IsNullOrEmpty(digits)) return null;

            // If the value doesn't contain a decimal point, add one based on digit count
            if (digits.Contains("."))
            {
                return double.TryParse(digits, out double v) ? v : (double?)null;
            }

            // Convert digits to decimal (e.g., "010" → 0.010, "005" → 0.005)
            string withDecimal = "0." + digits.PadLeft(3, '0');
            return double.TryParse(withDecimal, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double result) ? result : (double?)null;
        }
    }
}
