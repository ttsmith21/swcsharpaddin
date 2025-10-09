using System;

namespace NM.Core.Tubes
{
    /// <summary>
    /// Provides cutting parameters for tube processing.
    /// Units: inches and seconds. Feed rates are inches/min.
    /// </summary>
    public sealed class TubeCuttingParameterService
    {
        public sealed class CutParams
        {
            public double KerfIn { get; set; }
            public double PierceTimeSec { get; set; }
            public double CutSpeedInPerMin { get; set; }
        }

        /// <summary>
        /// Returns parameters based on material category and wall thickness, ported from VBA.
        /// materialCategory: "CarbonSteel", "StainlessSteel", "Aluminum" (case-insensitive).
        /// </summary>
        public CutParams Get(string materialCategory, double wallIn)
        {
            materialCategory = (materialCategory ?? string.Empty).Trim();
            var p = new CutParams();

            switch (materialCategory.ToLowerInvariant())
            {
                case "carbonsteel":
                    p.KerfIn = 0.02;
                    p.CutSpeedInPerMin = GetFeedRateCs(wallIn) * 0.85; // final 0.85 factor per VBA
                    p.PierceTimeSec = GetPierceTimeCs(wallIn);
                    break;
                case "stainlesssteel":
                    p.KerfIn = 0.02;
                    p.CutSpeedInPerMin = GetFeedRateSs(wallIn) * 0.85;
                    p.PierceTimeSec = GetPierceTimeSs(wallIn);
                    break;
                case "aluminum":
                    p.KerfIn = 0.03;
                    p.CutSpeedInPerMin = 0.0; // VBA returns 0 for AL feed
                    p.PierceTimeSec = 0.0;    // VBA returns 0 for AL pierce
                    break;
                default:
                    p.KerfIn = 0.02;
                    p.CutSpeedInPerMin = GetFeedRateCs(wallIn) * 0.85;
                    p.PierceTimeSec = GetPierceTimeCs(wallIn);
                    break;
            }
            return p;
        }

        private static double GetFeedRateCs(double t)
        {
            if (t <= 0.045) return 295;
            if (t <= 0.055) return 271;
            if (t <= 0.065) return 251;
            if (t <= 0.085) return 196;
            if (t <= 0.105) return 161;
            if (t <= 0.125) return 149;
            if (t <= 0.145) return 137;
            if (t <= 0.165) return 129;
            if (t <= 0.185) return 122;
            if (t <= 0.205) return 118;
            if (t <= 0.255) return 106;
            if (t <= 0.32) return 87;
            if (t <= 0.38) return 47;
            if (t <= 0.405) return 36;
            if (t <= 0.455) return 19;
            return 7;
        }

        private static double GetFeedRateSs(double t)
        {
            if (t <= 0.045) return 397;
            if (t <= 0.055) return 354;
            if (t <= 0.065) return 318;
            if (t <= 0.085) return 251;
            if (t <= 0.105) return 196;
            if (t <= 0.125) return 157;
            if (t <= 0.145) return 135;
            if (t <= 0.165) return 118;
            if (t <= 0.185) return 104;
            if (t <= 0.205) return 90;
            if (t <= 0.255) return 78;
            if (t <= 0.32) return 54;
            if (t <= 0.38) return 36;
            if (t <= 0.405) return 30;
            if (t <= 0.455) return 18;
            return 8;
        }

        private static double GetPierceTimeCs(double t)
        {
            if (t <= 0.085) return 0.05;
            if (t <= 0.105) return 0.07;
            if (t <= 0.125) return 0.10;
            if (t <= 0.145) return 0.20;
            if (t <= 0.165) return 0.30;
            if (t <= 0.185) return 0.40;
            if (t <= 0.205) return 0.50;
            if (t <= 0.255) return 0.70;
            if (t <= 0.32) return 2.80;
            // 0.32 .. 0.455 and beyond
            return 5.0;
        }

        private static double GetPierceTimeSs(double t)
        {
            if (t <= 0.085) return 0.05;
            if (t <= 0.105) return 0.07;
            if (t <= 0.125) return 0.08;
            if (t <= 0.145) return 0.45;
            if (t <= 0.165) return 0.60;
            if (t <= 0.205) return 0.60;
            if (t <= 0.255) return 2.00;
            if (t <= 0.32) return 3.00;
            // 0.32 .. 0.455 and beyond
            return 5.0;
        }
    }
}
