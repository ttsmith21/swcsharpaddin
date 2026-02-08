using System;
using System.Collections.Generic;

namespace NM.Core.Manufacturing
{
    /// <summary>
    /// Result of hole/cutout-to-bend-line proximity analysis.
    /// Detects features in the deformation zone that will distort during bending.
    /// </summary>
    public sealed class HoleNearBendResult
    {
        public int ViolationCount { get; set; }
        public List<HoleNearBendViolation> Violations { get; } = new List<HoleNearBendViolation>();
        public bool HasViolations => ViolationCount > 0;
    }

    /// <summary>
    /// A single hole/cutout that is too close to a bend line.
    /// </summary>
    public sealed class HoleNearBendViolation
    {
        public int HoleIndex { get; set; }
        public int BendLineIndex { get; set; }
        public double ActualDistanceIn { get; set; }
        public double RequiredDistanceIn { get; set; }
        public string Description { get; set; }
    }

    /// <summary>
    /// Material-specific multiplier for hole-to-bend minimum distance.
    /// Northern's rule: SS = 5T + R, CRS/AL = 4T + R.
    /// </summary>
    public static class BendClearanceLookup
    {
        /// <summary>
        /// Returns the thickness multiplier for hole-to-bend distance based on material.
        /// SS 304/316 = 5, everything else (CRS, AL) = 4.
        /// </summary>
        public static int GetMultiplier(string material)
        {
            if (string.IsNullOrWhiteSpace(material))
                return 4; // default to CRS/AL rule

            var m = material.ToUpperInvariant();

            // Stainless steel: 5T + R
            if (m.Contains("304") || m.Contains("316") || m.Contains("309") ||
                m.Contains("STAINLESS") || m.Contains("SS "))
                return 5;

            // CRS, aluminum, everything else: 4T + R
            return 4;
        }

        /// <summary>
        /// Computes the minimum distance (inches) from any cut feature edge to a bend line.
        /// Formula: multiplier * T + R
        /// </summary>
        public static double ComputeMinDistance(string material, double thicknessIn, double bendRadiusIn)
        {
            int mult = GetMultiplier(material);
            return (mult * thicknessIn) + bendRadiusIn;
        }
    }
}
