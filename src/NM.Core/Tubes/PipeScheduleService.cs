using System;
using System.Collections.Generic;
using System.Globalization;

namespace NM.Core.Tubes
{
    /// <summary>
    /// Pipe schedule resolution utilities.
    /// - Tolerant OD/wall -> (NPS text, schedule code) resolver matching VBA PipeDiam.
    /// - Optional legacy NPS+schedule DB (limited).
    /// Units: inches.
    /// </summary>
    public sealed class PipeScheduleService
    {
        public sealed class PipeSpec
        {
            public double OutsideDiameterIn { get; set; }
            public double WallThicknessIn { get; set; }
            public double InsideDiameterIn => System.Math.Max(0.0, OutsideDiameterIn - 2.0 * WallThicknessIn);
        }

        private readonly Dictionary<(double nps, string sch), PipeSpec> _legacyDb = new Dictionary<(double, string), PipeSpec>();

        // OD -> { wall -> scheduleCode }
        private static readonly Dictionary<double, Dictionary<double, string>> _odMap = new Dictionary<double, Dictionary<double, string>>
        {
            { 0.405, new Dictionary<double, string> { {0.049,"10"},{0.068,"40"},{0.095,"80"} } },
            { 0.540, new Dictionary<double, string> { {0.065,"10"},{0.088,"40"},{0.119,"80"} } },
            { 0.675, new Dictionary<double, string> { {0.065,"10"},{0.091,"40"},{0.126,"80"} } },
            { 0.840, new Dictionary<double, string> { {0.065,"5"},{0.083,"10"},{0.109,"40"},{0.147,"80"},{0.187,"160"},{0.294,"XX"} } },
            { 1.050, new Dictionary<double, string> { {0.065,"5"},{0.083,"10"},{0.113,"40"},{0.154,"80"},{0.218,"160"},{0.308,"XX"} } },
            { 1.315, new Dictionary<double, string> { {0.065,"5"},{0.109,"10"},{0.133,"40"},{0.179,"80"},{0.250,"160"},{0.358,"XX"} } },
            { 1.660, new Dictionary<double, string> { {0.065,"5"},{0.109,"10"},{0.140,"40"},{0.191,"80"},{0.250,"160"},{0.382,"XX"} } },
            { 1.900, new Dictionary<double, string> { {0.065,"5"},{0.109,"10"},{0.145,"40"},{0.200,"80"},{0.281,"160"},{0.400,"XX"} } },
            { 2.375, new Dictionary<double, string> { {0.065,"5"},{0.120,"10"},{0.154,"40"},{0.218,"80"},{0.344,"160"},{0.436,"XX"} } },
            { 2.875, new Dictionary<double, string> { {0.083,"5"},{0.120,"10"},{0.203,"40"},{0.276,"80"},{0.375,"160"},{0.552,"XX"} } },
            { 3.500, new Dictionary<double, string> { {0.083,"5"},{0.120,"10"},{0.216,"40"},{0.300,"80"},{0.438,"160"},{0.600,"XX"} } },
            { 4.000, new Dictionary<double, string> { {0.083,"5"},{0.120,"10"},{0.226,"40"},{0.318,"80"},{0.636,"XX"} } },
            { 4.500, new Dictionary<double, string> { {0.083,"5"},{0.120,"10"},{0.237,"40"},{0.337,"80"},{0.438,"120"},{0.531,"160"},{0.674,"XX"} } },
            { 5.000, new Dictionary<double, string> { {0.247,"STD"},{0.120,"XX"} } },
            { 5.563, new Dictionary<double, string> { {0.109,"5"},{0.134,"10"},{0.258,"40"},{0.375,"80"},{0.500,"120"},{0.625,"160"},{0.750,"XX"} } },
            { 6.625, new Dictionary<double, string> { {0.109,"5"},{0.134,"10"},{0.280,"40"},{0.432,"80"},{0.562,"120"},{0.718,"160"},{0.864,"XX"} } },
            { 8.625, new Dictionary<double, string> { {0.109,"5"},{0.148,"10"},{0.322,"40"},{0.500,"80"},{0.718,"120"},{0.906,"160"},{0.875,"XX"} } },
            {10.750, new Dictionary<double, string> { {0.134,"5"},{0.165,"10"},{0.365,"40"},{0.500,"80S"},{0.593,"80"},{0.843,"120"},{1.125,"160"} } },
            {12.750, new Dictionary<double, string> { {0.156,"5"},{0.180,"10"},{0.375,"40S"},{0.406,"40"},{0.500,"80S"},{0.687,"80"},{1.000,"120"},{1.312,"160"} } },
            {14.000, new Dictionary<double, string> { {0.156,"5"},{0.188,"10S"},{0.250,"10"},{0.375,"40S"},{0.437,"40"},{0.500,"80S"},{0.750,"80"},{1.093,"120"},{1.406,"160"} } },
            {16.000, new Dictionary<double, string> { {0.156,"5"},{0.188,"10S"},{0.250,"10"},{0.375,"40S"},{0.843,"80"},{1.218,"120"},{1.437,"160"} } },
            {18.000, new Dictionary<double, string> { {0.165,"5"},{0.188,"10S"},{0.250,"10"},{0.375,"40S"},{0.562,"40"},{0.500,"80S"},{0.937,"80"},{1.375,"120"},{1.781,"160"} } },
            {20.000, new Dictionary<double, string> { {0.188,"5"},{0.218,"10S"},{0.250,"10"},{0.375,"40S"},{0.593,"40"},{0.500,"80S"},{1.031,"80"},{1.500,"120"},{1.968,"160"} } },
            {24.000, new Dictionary<double, string> { {0.218,"5"},{0.250,"10"},{0.375,"40S"},{0.687,"40"},{0.500,"80S"},{1.218,"80"},{1.812,"120"},{2.343,"160"} } },
        };

        public PipeScheduleService()
        {
            // Legacy small subset (kept for backwards TryGet compatibility)
            AddLegacy(0.5, "40", 0.840, 0.109);
            AddLegacy(1.0, "40", 1.315, 0.133);
            AddLegacy(2.0, "40", 2.375, 0.154);
            AddLegacy(3.0, "40", 3.500, 0.216);
            AddLegacy(4.0, "40", 4.500, 0.237);
            AddLegacy(6.0, "40", 6.625, 0.280);
        }

        private void AddLegacy(double nps, string schedule, double od, double wall)
        {
            _legacyDb[(nps, schedule.ToUpperInvariant())] = new PipeSpec { OutsideDiameterIn = od, WallThicknessIn = wall };
        }

        public bool TryGet(double npsInches, string schedule, out PipeSpec spec)
        {
            spec = null;
            if (string.IsNullOrWhiteSpace(schedule)) return false;
            return _legacyDb.TryGetValue((npsInches, schedule.ToUpperInvariant()), out spec);
        }

        /// <summary>
        /// Resolves a pipe schedule by measured OD and wall thickness within small tolerances.
        /// Returns NPS label text like 1" or .5" (before SCH) and the schedule code (e.g., 40, 80S, XX).
        /// Applies a stainless special-case for 16" OD with 0.500 wall -> 80S, else 40.
        /// </summary>
        public bool TryResolveByOdAndWall(double odIn, double wallIn, string materialCategory, out string npsText, out string scheduleCode)
        {
            const double odTol = 0.010; // inches
            const double wallTol = 0.005; // inches
            npsText = string.Empty; scheduleCode = string.Empty;

            double bestOd = 0.0; string bestSched = null;
            foreach (var kv in _odMap)
            {
                if (System.Math.Abs(kv.Key - odIn) > odTol) continue;
                foreach (var wt in kv.Value)
                {
                    if (System.Math.Abs(wt.Key - wallIn) <= wallTol)
                    {
                        bestOd = kv.Key; bestSched = wt.Value; break;
                    }
                }
                if (bestSched != null) break;
            }

            if (bestSched == null && System.Math.Abs(odIn - 16.000) <= odTol && System.Math.Abs(wallIn - 0.500) <= wallTol)
            {
                bestOd = 16.000;
                bestSched = (materialCategory ?? string.Empty).Equals("StainlessSteel", System.StringComparison.OrdinalIgnoreCase) ? "80S" : "40";
            }

            if (bestSched == null) return false;

            npsText = MapOdToNpsText(bestOd);
            scheduleCode = bestSched;
            return true;
        }

        private static string MapOdToNpsText(double od)
        {
            if (System.Math.Abs(od - 0.405) < 1e-3) return ".125\"";
            if (System.Math.Abs(od - 0.540) < 1e-3) return ".25\"";
            if (System.Math.Abs(od - 0.675) < 1e-3) return ".375\"";
            if (System.Math.Abs(od - 0.840) < 1e-3) return ".5\"";
            if (System.Math.Abs(od - 1.050) < 1e-3) return ".75\"";
            if (System.Math.Abs(od - 1.315) < 1e-3) return "1\"";
            if (System.Math.Abs(od - 1.660) < 1e-3) return "1.25\"";
            if (System.Math.Abs(od - 1.900) < 1e-3) return "1.5\"";
            if (System.Math.Abs(od - 2.375) < 1e-3) return "2\"";
            if (System.Math.Abs(od - 2.875) < 1e-3) return "2.5\"";
            if (System.Math.Abs(od - 3.500) < 1e-3) return "3\"";
            if (System.Math.Abs(od - 4.000) < 1e-3) return "3.5\"";
            if (System.Math.Abs(od - 4.500) < 1e-3) return "4\"";
            if (System.Math.Abs(od - 5.563) < 1e-3) return "5\"";
            if (System.Math.Abs(od - 6.625) < 1e-3) return "6\"";
            if (System.Math.Abs(od - 8.625) < 1e-3) return "8\"";
            if (System.Math.Abs(od - 10.750) < 1e-3) return "10\"";
            if (System.Math.Abs(od - 12.750) < 1e-3) return "12\"";
            if (System.Math.Abs(od - 14.000) < 1e-3) return "14\"";
            if (System.Math.Abs(od - 16.000) < 1e-3) return "16\"";
            if (System.Math.Abs(od - 18.000) < 1e-3) return "18\"";
            if (System.Math.Abs(od - 20.000) < 1e-3) return "20\"";
            if (System.Math.Abs(od - 24.000) < 1e-3) return "24\"";
            return od.ToString("0.###\"", CultureInfo.InvariantCulture);
        }
    }
}
