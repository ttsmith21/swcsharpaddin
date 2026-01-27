using System;
using System.Globalization;

namespace NM.Core.Tubes
{
    /// <summary>
    /// Generates OptiMaterial-like codes for tubes based on OD, wall, and material.
    /// </summary>
    public static class TubeMaterialCodeGenerator
    {
        /// <summary>
        /// Example code format: TUBE-{Material}-{OD*1000:0}-{Wall*1000:0}
        /// OD and Wall in inches.
        /// </summary>
        public static string Generate(string material, double odIn, double wallIn)
        {
            material = (material ?? "").Trim().ToUpperInvariant();
            if (string.IsNullOrEmpty(material)) material = "UNKNOWN";
            int odMil = (int)Math.Round(odIn * 1000.0);
            int wallMil = (int)Math.Round(wallIn * 1000.0);
            return $"TUBE-{material}-{odMil}-{wallMil}";
        }
    }
}
