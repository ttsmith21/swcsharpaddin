namespace NM.Core.Manufacturing
{
    /// <summary>
    /// Evaluates whether the default nesting efficiency (80%) overestimates material usage
    /// by comparing the part's mass to its bounding-box blank weight.
    /// When the part fills less than the default efficiency of its bounding rectangle,
    /// recommends switching to the L×W (manual) weight calculation mode.
    /// </summary>
    public static class NestingEfficiencyEvaluator
    {
        /// <summary>
        /// Border added to each bounding-box dimension (inches) to account for
        /// kerf width and inter-part spacing on the sheet.
        /// </summary>
        public static double BorderInches { get; set; } = 0.25;

        private const double DEFAULT_NEST_EFFICIENCY = 80.0; // percent
        private const double OVERRIDE_THRESHOLD = 50.0; // percent — only flag parts below this

        public sealed class EvalResult
        {
            /// <summary>True when the evaluator recommends overriding to L×W mode.</summary>
            public bool ShouldOverride { get; set; }

            /// <summary>Calculated bounding-box nesting efficiency (0-100%).</summary>
            public double BBoxEfficiencyPercent { get; set; }

            /// <summary>Blank length including border (inches).</summary>
            public double BlankLengthIn { get; set; }

            /// <summary>Blank width including border (inches).</summary>
            public double BlankWidthIn { get; set; }

            /// <summary>Calculated blank weight (lb).</summary>
            public double BlankWeightLb { get; set; }

            /// <summary>Human-readable explanation of the decision.</summary>
            public string Reason { get; set; }
        }

        /// <summary>
        /// Evaluates whether the bounding-box efficiency is below the override threshold
        /// (50%) and the user has not customised the weight-calc settings.
        /// </summary>
        /// <param name="partMassLb">Finished part mass in pounds.</param>
        /// <param name="bboxLengthIn">Flat-pattern bounding-box length in inches (already ≥ width).</param>
        /// <param name="bboxWidthIn">Flat-pattern bounding-box width in inches.</param>
        /// <param name="thicknessIn">Sheet thickness in inches.</param>
        /// <param name="densityLbPerIn3">Material density in lb/in³.</param>
        /// <param name="currentCalcMode">"0" = efficiency mode, "1" = manual L×W mode.</param>
        /// <param name="currentNestEfficiency">Current nesting efficiency percent (e.g. 80).</param>
        public static EvalResult Evaluate(
            double partMassLb,
            double bboxLengthIn,
            double bboxWidthIn,
            double thicknessIn,
            double densityLbPerIn3,
            string currentCalcMode,
            double currentNestEfficiency)
        {
            var result = new EvalResult();

            // Guard: need valid inputs
            if (partMassLb <= 0 || bboxLengthIn <= 0 || bboxWidthIn <= 0 ||
                thicknessIn <= 0 || densityLbPerIn3 <= 0)
            {
                result.Reason = "Insufficient data for nesting efficiency evaluation";
                return result;
            }

            // Only evaluate if user is on defaults (efficiency mode with 80%)
            bool isEfficiencyMode = (currentCalcMode ?? "0") == "0";
            bool isDefaultEfficiency = System.Math.Abs(currentNestEfficiency - DEFAULT_NEST_EFFICIENCY) < 0.01;

            if (!isEfficiencyMode || !isDefaultEfficiency)
            {
                result.Reason = "User has customised weight calculation settings; skipping auto-evaluation";
                return result;
            }

            // Calculate blank dimensions with border
            double blankL = bboxLengthIn + BorderInches;
            double blankW = bboxWidthIn + BorderInches;
            double blankWeightLb = blankL * blankW * thicknessIn * densityLbPerIn3;

            result.BlankLengthIn = blankL;
            result.BlankWidthIn = blankW;
            result.BlankWeightLb = blankWeightLb;

            if (blankWeightLb <= 0)
            {
                result.Reason = "Blank weight calculation returned zero";
                return result;
            }

            double bboxEfficiency = (partMassLb / blankWeightLb) * 100.0;
            result.BBoxEfficiencyPercent = bboxEfficiency;

            if (bboxEfficiency < OVERRIDE_THRESHOLD)
            {
                result.ShouldOverride = true;
                result.Reason = $"Bounding-box efficiency {bboxEfficiency:F1}% is below override threshold {OVERRIDE_THRESHOLD:F0}%. " +
                                $"Blank size: {blankL:F2}\" x {blankW:F2}\". " +
                                $"Switching to L×W mode to avoid underestimating material cost.";
            }
            else
            {
                result.Reason = $"Bounding-box efficiency {bboxEfficiency:F1}% meets or exceeds override threshold {OVERRIDE_THRESHOLD:F0}%. No override needed.";
            }

            return result;
        }
    }
}
