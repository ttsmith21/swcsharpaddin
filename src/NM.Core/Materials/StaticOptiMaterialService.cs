using System;
using System.Collections.Generic;
using System.Globalization;
using NM.Core.DataModel;

namespace NM.Core.Materials
{
    /// <summary>
    /// Generates OptiMaterial codes without Excel data using hardcoded gauge tables.
    /// Fallback for when Laser2022v4.xlsx is unavailable (e.g., QA headless mode).
    ///
    /// Code patterns:
    ///   Sheet metal: S.{material}{gauge}  e.g., S.304L14GA, S.304L.25IN
    ///   Pipe:        P.{material}{nps}SCH{schedule}  e.g., P.304L12"SCH40S
    ///   Tube:        T.{material}{dims}  e.g., T.304L2"X1"X.060"
    ///   Angle:       A.{material}{dims}  e.g., A.304L2"X2"X.125"
    ///   Round bar:   R.{material}{dia}  e.g., R.304L1"
    /// </summary>
    public static class StaticOptiMaterialService
    {
        // Stainless steel gauge table: thickness (inches) -> gauge name
        // Standard stainless steel gauges (slightly different from carbon steel)
        private static readonly SortedList<double, string> _ssGaugeTable = new SortedList<double, string>
        {
            { 0.0187, "26GA" },
            { 0.0217, "24GA" },
            { 0.0250, "23GA" },
            { 0.0299, "22GA" },
            { 0.0359, "20GA" },
            { 0.0478, "18GA" },
            { 0.0598, "16GA" },
            { 0.0747, "14GA" },
            { 0.1046, "12GA" },
            { 0.1196, "11GA" },
            { 0.1345, "10GA" },
        };

        // Common fractional/decimal thicknesses expressed as label
        private static readonly SortedList<double, string> _thicknessLabels = new SortedList<double, string>
        {
            { 0.1875, "3/16" },
            { 0.250,  ".25IN" },
            { 0.3125, "5/16" },
            { 0.375,  "3/8" },
            { 0.500,  ".5IN" },
            { 0.625,  "5/8" },
            { 0.750,  "3/4" },
            { 1.000,  "1IN" },
            { 1.250,  "1.25IN" },
            { 1.500,  "1.5IN" },
            { 2.000,  "2IN" },
        };

        private const double GAUGE_TOLERANCE = 0.008; // inches

        /// <summary>
        /// Resolve OptiMaterial code from part data.
        /// Returns null if insufficient data to generate a code.
        /// </summary>
        public static string Resolve(PartData pd)
        {
            if (pd == null) return null;

            string materialCode = MaterialCodeMapper.ToShortCode(pd.Material);
            if (string.IsNullOrWhiteSpace(materialCode)) return null;

            // Sheet metal parts
            if (pd.Classification == PartType.SheetMetal ||
                (pd.Sheet != null && pd.Sheet.IsSheetMetal))
            {
                return ResolveSheetMetal(pd, materialCode);
            }

            // Tube/pipe/structural parts
            if (pd.Classification == PartType.Tube && pd.Tube != null && pd.Tube.IsTube)
            {
                return ResolveTube(pd, materialCode);
            }

            return null;
        }

        private static string ResolveSheetMetal(PartData pd, string materialCode)
        {
            const double M_TO_IN = 39.3701;
            double thicknessIn = pd.Thickness_m > 0 ? pd.Thickness_m * M_TO_IN : 0;
            if (thicknessIn <= 0) return null;

            string thicknessLabel = ResolveThicknessLabel(thicknessIn);
            return "S." + materialCode + thicknessLabel;
        }

        private static string ResolveTube(PartData pd, string materialCode)
        {
            const double M_TO_IN = 39.3701;

            string shape = (pd.Tube.TubeShape ?? "").Trim();

            // Pipe (round with NPS/schedule)
            if (shape.Equals("Round", StringComparison.OrdinalIgnoreCase) ||
                string.IsNullOrEmpty(shape))
            {
                // If we have NPS and schedule, it's a pipe
                if (!string.IsNullOrEmpty(pd.Tube.NpsText) && !string.IsNullOrEmpty(pd.Tube.ScheduleCode))
                {
                    return "P." + materialCode + pd.Tube.NpsText + "SCH" + pd.Tube.ScheduleCode;
                }

                // Round tube without NPS - use OD and wall
                double odIn = pd.Tube.OD_m > 0 ? pd.Tube.OD_m * M_TO_IN : 0;
                double wallIn = pd.Tube.Wall_m > 0 ? pd.Tube.Wall_m * M_TO_IN : 0;
                if (odIn > 0 && wallIn > 0)
                {
                    return "T." + materialCode + FormatDim(odIn) + "ODX" + FormatDim(wallIn);
                }
            }

            // Angle
            if (shape.Equals("Angle", StringComparison.OrdinalIgnoreCase))
            {
                double odIn = pd.Tube.OD_m > 0 ? pd.Tube.OD_m * M_TO_IN : 0;
                double idIn = pd.Tube.ID_m > 0 ? pd.Tube.ID_m * M_TO_IN : 0;
                double wallIn = pd.Tube.Wall_m > 0 ? pd.Tube.Wall_m * M_TO_IN : 0;
                if (odIn > 0 && wallIn > 0)
                {
                    // VBA format: A.304L2"X2"X.125" (leg1 x leg2 x wall)
                    string leg2 = idIn > 0 ? FormatDim(idIn) : FormatDim(odIn);
                    return "A." + materialCode + FormatDim(odIn) + "X" + leg2 + "X" + FormatDim(wallIn);
                }
            }

            // Round bar (no wall thickness = solid)
            if (shape.Equals("Round Bar", StringComparison.OrdinalIgnoreCase))
            {
                double odIn = pd.Tube.OD_m > 0 ? pd.Tube.OD_m * M_TO_IN : 0;
                if (odIn > 0)
                {
                    return "R." + materialCode + FormatDim(odIn);
                }
            }

            // Square tube
            if (shape.Equals("Square", StringComparison.OrdinalIgnoreCase))
            {
                double odIn = pd.Tube.OD_m > 0 ? pd.Tube.OD_m * M_TO_IN : 0;
                double wallIn = pd.Tube.Wall_m > 0 ? pd.Tube.Wall_m * M_TO_IN : 0;
                if (odIn > 0 && wallIn > 0)
                {
                    return "T." + materialCode + FormatDim(odIn) + "SQX" + FormatDim(wallIn);
                }
            }

            // Rectangular tube
            if (shape.Equals("Rectangle", StringComparison.OrdinalIgnoreCase))
            {
                double odIn = pd.Tube.OD_m > 0 ? pd.Tube.OD_m * M_TO_IN : 0;
                double idIn = pd.Tube.ID_m > 0 ? pd.Tube.ID_m * M_TO_IN : 0;
                double wallIn = pd.Tube.Wall_m > 0 ? pd.Tube.Wall_m * M_TO_IN : 0;
                if (odIn > 0 && wallIn > 0)
                {
                    // For rectangular: OD is long side, need short side
                    // VBA: T.304L2"X1"X.060"
                    string shortSide = idIn > 0 ? FormatDim(idIn) : FormatDim(odIn);
                    return "T." + materialCode + FormatDim(odIn) + "X" + shortSide + "X" + FormatDim(wallIn);
                }
            }

            // Channel / I-Beam / other structural
            if (shape.Equals("Channel", StringComparison.OrdinalIgnoreCase) ||
                shape.Equals("I-Beam", StringComparison.OrdinalIgnoreCase) ||
                shape.Equals("Beam", StringComparison.OrdinalIgnoreCase))
            {
                double odIn = pd.Tube.OD_m > 0 ? pd.Tube.OD_m * M_TO_IN : 0;
                double idIn = pd.Tube.ID_m > 0 ? pd.Tube.ID_m * M_TO_IN : 0;
                double wallIn = pd.Tube.Wall_m > 0 ? pd.Tube.Wall_m * M_TO_IN : 0;
                if (odIn > 0 && wallIn > 0)
                {
                    // VBA uses C. prefix for channels, T. for I-beams
                    string prefix = shape.Equals("Channel", StringComparison.OrdinalIgnoreCase) ? "C." : "T.";
                    string dim2 = idIn > 0 ? FormatDim(idIn) : FormatDim(odIn);
                    return prefix + materialCode + FormatDim(odIn) + "X" + dim2 + "X" + FormatDim(wallIn);
                }
            }

            return null;
        }

        /// <summary>
        /// Resolve thickness to a gauge label or decimal inches label.
        /// </summary>
        internal static string ResolveThicknessLabel(double thicknessIn)
        {
            // Try standard gauges first
            foreach (var kvp in _ssGaugeTable)
            {
                if (Math.Abs(kvp.Key - thicknessIn) <= GAUGE_TOLERANCE)
                    return kvp.Value;
            }

            // Try common fractional/decimal thicknesses
            foreach (var kvp in _thicknessLabels)
            {
                if (Math.Abs(kvp.Key - thicknessIn) <= GAUGE_TOLERANCE)
                    return kvp.Value;
            }

            // Fall back to decimal inches formatting
            // Use VBA convention: ".187IN" for sub-inch, "1.5IN" for over-inch
            if (thicknessIn < 1.0)
            {
                string formatted = thicknessIn.ToString("G4", CultureInfo.InvariantCulture);
                // Ensure leading dot for values < 1
                if (!formatted.StartsWith("0.") && !formatted.StartsWith("."))
                    formatted = "." + formatted;
                else if (formatted.StartsWith("0."))
                    formatted = formatted.Substring(1); // "0.187" -> ".187"
                return formatted + "IN";
            }

            return thicknessIn.ToString("G4", CultureInfo.InvariantCulture) + "IN";
        }

        /// <summary>
        /// Format a dimension in inches with trailing quote mark.
        /// E.g., 2.0 -> "2\"", 0.125 -> ".125\"", 12.75 -> "12.75\""
        /// </summary>
        private static string FormatDim(double inches)
        {
            // Format to match VBA conventions
            if (inches == Math.Floor(inches))
                return ((int)inches).ToString() + "\"";

            string s;
            if (inches < 1.0)
            {
                // Sub-inch: 3 decimal places to match VBA gauge conventions (.060", .250", .170")
                s = inches.ToString("F3", CultureInfo.InvariantCulture);
                // VBA convention: trim double trailing zeros (.500->.5) but keep single (.060 stays)
                if (s.EndsWith("00"))
                    s = s.Substring(0, s.Length - 2);
                if (s.StartsWith("0."))
                    s = s.Substring(1); // "0.060" -> ".060"
            }
            else
            {
                // >= 1 inch: significant digits (5.9", 3.94", 12.75")
                s = inches.ToString("G4", CultureInfo.InvariantCulture);
            }
            return s + "\"";
        }
    }
}
