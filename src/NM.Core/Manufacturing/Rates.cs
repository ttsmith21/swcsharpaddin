namespace NM.Core.Manufacturing
{
    internal static class Rates
    {
        // Densities (lb/in^3)
        public const double Density304_316 = 0.289;
        public const double DensityA36 = 0.284;
        public const double DensityAl = 0.098; // 6061/5052 approx

        public static double GetDensityLbPerIn3(string material)
        {
            var m = (material ?? string.Empty).ToUpperInvariant();
            if (m.Contains("304") || m.Contains("316")) return Density304_316;
            if (m.Contains("A36") || m.Contains("CS") || m.Contains("CARBON")) return DensityA36;
            if (m.Contains("6061") || m.Contains("5052") || m.Contains("AL")) return DensityAl;
            return Density304_316; // default conservative
        }

        // Laser rates (rough; v1 placeholders)
        public static double GetLaserIPM(string material)
        {
            var m = (material ?? string.Empty).ToUpperInvariant();
            if (m.Contains("AL")) return 100; // aluminum faster
            return 60; // SS/CS default
        }

        public const double DefaultPierceSeconds = 0.5; // per pierce

        // Brake per-bend seconds by thickness (very rough tiers)
        public static int GetBrakeSecondsPerBend(double thicknessIn)
        {
            if (thicknessIn <= 0.075) return 10;
            if (thicknessIn <= 0.135) return 30;
            if (thicknessIn <= 0.250) return 45;
            return 60;
        }
    }
}
