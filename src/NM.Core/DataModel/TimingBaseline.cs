using System;
using System.Collections.Generic;
using System.IO;

namespace NM.Core.DataModel
{
    /// <summary>
    /// Represents baseline timing data for regression detection.
    /// Stored in tests/timing-baseline.json.
    /// </summary>
    public sealed class TimingBaseline
    {
        public string CreatedAt { get; set; }
        public string RunId { get; set; }
        public int FileCount { get; set; }
        public double TotalElapsedMs { get; set; }
        public Dictionary<string, BaselineTimerEntry> Timers { get; set; } = new Dictionary<string, BaselineTimerEntry>();

        /// <summary>
        /// Default path for baseline file.
        /// </summary>
        public static string DefaultPath
        {
            get
            {
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                // Navigate from bin/Debug to tests/
                return Path.Combine(baseDir, "..", "..", "tests", "timing-baseline.json");
            }
        }
    }

    /// <summary>
    /// Baseline data for a single timer category.
    /// </summary>
    public sealed class BaselineTimerEntry
    {
        public double AvgMs { get; set; }
        public double MaxMs { get; set; }
        public int Count { get; set; }
    }

    /// <summary>
    /// Compares current timing against baseline and detects regressions.
    /// </summary>
    public sealed class TimingRegressionDetector
    {
        /// <summary>
        /// Percentage threshold above which a slowdown is flagged as a regression.
        /// Default: 20% (flag if current is >20% slower than baseline).
        /// </summary>
        public double RegressionThresholdPercent { get; set; } = 20.0;

        /// <summary>
        /// Detects regressions by comparing current timing against baseline.
        /// Returns list of regressions (slowdowns exceeding threshold).
        /// </summary>
        public List<TimingRegression> DetectRegressions(
            Dictionary<string, QATimingSummary> current,
            TimingBaseline baseline)
        {
            var regressions = new List<TimingRegression>();
            if (baseline?.Timers == null || current == null) return regressions;

            foreach (var kvp in current)
            {
                if (!baseline.Timers.TryGetValue(kvp.Key, out var baselineEntry)) continue;

                double baselineAvg = baselineEntry.AvgMs;
                double currentAvg = kvp.Value.AvgMs;

                // Skip if baseline is essentially zero
                if (baselineAvg <= 0.1) continue;

                double changePercent = ((currentAvg - baselineAvg) / baselineAvg) * 100.0;

                if (changePercent > RegressionThresholdPercent)
                {
                    regressions.Add(new TimingRegression
                    {
                        TimerName = kvp.Key,
                        BaselineAvgMs = baselineAvg,
                        CurrentAvgMs = currentAvg,
                        ChangePercent = changePercent,
                        Severity = changePercent > 50 ? "Critical" : "Warning"
                    });
                }
            }

            return regressions;
        }
    }

    /// <summary>
    /// Represents a detected timing regression.
    /// </summary>
    public sealed class TimingRegression
    {
        public string TimerName { get; set; }
        public double BaselineAvgMs { get; set; }
        public double CurrentAvgMs { get; set; }
        public double ChangePercent { get; set; }
        public string Severity { get; set; }  // "Warning" or "Critical"

        public override string ToString()
        {
            return $"[{Severity}] {TimerName}: {BaselineAvgMs:F1}ms -> {CurrentAvgMs:F1}ms (+{ChangePercent:F0}%)";
        }
    }
}
