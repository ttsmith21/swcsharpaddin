using System;

namespace NM.Core.Tubes
{
    /// <summary>
    /// Quick check to classify round bars vs tubes based on OD and ID.
    /// </summary>
    public static class RoundBarValidator
    {
        /// <summary>
        /// Returns true if the inside diameter is effectively zero (solid bar), using a small tolerance.
        /// Units: inches.
        /// </summary>
        public static bool IsRoundBar(double outsideDiameterIn, double insideDiameterIn, double tolIn = 1e-3)
        {
            return insideDiameterIn <= tolIn;
        }
    }
}
