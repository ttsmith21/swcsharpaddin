using System;
using System.Globalization;

namespace NM.Core.Tubes
{
    /// <summary>
    /// Work-center time estimators for tube parts (port of legacy VBA thresholds).
    /// All times returned in hours.
    /// </summary>
    public static class TubeWorkCenterRules
    {
        public sealed class F325Result
        {
            public string Code { get; set; } // usually "1"
            public double SetupHours { get; set; }
            public double RunHours { get; set; }
            public bool RequiresPressBrake { get; set; } // whether to apply F140
        }

        public sealed class F140Result
        {
            public double SetupHours { get; set; } // fixed 0.2 hours
            public double RunHours { get; set; }   // 0.08 (light) or 0.25 (heavy)
        }

        public sealed class F210Result
        {
            public double SetupHours { get; set; } // 0.03 hours
            public double RunHours { get; set; }   // length / DeburRate
        }

        /// <summary>
        /// Computes Roll Form (F325) per VBA:
        /// - F325_R = (weightLb * 5 / 3600) + (5 / 60)
        /// - F325_S thresholds:
        ///   * weight &lt; 40 -> 0.25 (no press brake)
        ///   * 40-150: 0.375; press brake only if thicknessIn &gt;= 0.165
        ///   * &gt;=150: 0.75; press brake only if thicknessIn &gt;= 0.165
        /// Thickness &lt; 0.165 alone does not trigger press brake.
        /// </summary>
        public static F325Result ComputeF325(double weightLb, double thicknessIn)
        {
            var res = new F325Result { Code = "1", RequiresPressBrake = false };
            // Run hours per VBA formula
            res.RunHours = (weightLb * (5.0 / 3600.0)) + (5.0 / 60.0);

            if (weightLb < 40.0)
            {
                res.SetupHours = 0.25;
                res.RequiresPressBrake = false;
            }
            else if (weightLb < 150.0)
            {
                res.SetupHours = 0.375;
                res.RequiresPressBrake = thicknessIn >= 0.165;
            }
            else
            {
                res.SetupHours = 0.75;
                res.RequiresPressBrake = thicknessIn >= 0.165;
            }
            return res;
        }

        /// <summary>
        /// Computes Press Brake (F140) times when required (thickness &gt;= 0.165):
        /// - Setup = 0.2
        /// - Run = 0.08 for 40-150 lbs; 0.25 for &gt;=150 lbs; (no press brake for weight &lt; 40)
        /// </summary>
        public static F140Result ComputeF140(double weightLb, double thicknessIn)
        {
            var res = new F140Result { SetupHours = 0.20, RunHours = 0.0 };
            if (thicknessIn < 0.165) return res; // not applied; Run=0

            if (weightLb < 40.0)
            {
                res.RunHours = 0.0; // not applied for very light per thresholds
            }
            else if (weightLb < 150.0)
            {
                res.RunHours = 0.08;
            }
            else
            {
                res.RunHours = 0.25;
            }
            return res;
        }

        /// <summary>
        /// Computes Deburr (F210) time: F210_R = lengthIn / DeburRate; F210_S = 0.03.
        /// DeburRate default: 3600 in/hr (60 in/min).
        /// </summary>
        public static F210Result ComputeF210(double lengthIn, double deburRateInPerHour = 3600.0)
        {
            if (deburRateInPerHour <= 0) deburRateInPerHour = 3600.0;
            var res = new F210Result
            {
                SetupHours = 0.03,
                RunHours = lengthIn / deburRateInPerHour
            };
            return res;
        }
    }
}
