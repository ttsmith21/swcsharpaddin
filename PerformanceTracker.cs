using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace NM.Core
{
    /// <summary>
    /// Comprehensive performance tracking service. Thread-safe and integrates with ErrorHandler.
    /// </summary>
    public sealed class PerformanceTracker : IDisposable
    {
        private static readonly Lazy<PerformanceTracker> _instance = new Lazy<PerformanceTracker>(() => new PerformanceTracker());
        public static PerformanceTracker Instance => _instance.Value;

        private readonly ConcurrentDictionary<string, List<TimerInfo>> _timers = new ConcurrentDictionary<string, List<TimerInfo>>(StringComparer.OrdinalIgnoreCase);
        private readonly ConcurrentDictionary<string, Stopwatch> _running = new ConcurrentDictionary<string, Stopwatch>(StringComparer.OrdinalIgnoreCase);
        private readonly object _listGate = new object();
        private volatile bool _disposed;

        private PerformanceTracker() { }

        public bool IsEnabled => Configuration.Logging.EnablePerformanceMonitoring && !Configuration.Logging.IsProductionMode;

        /// <summary>
        /// Log a mode-transition event that is always captured when logging is on.
        /// Use for significant state changes (scope enter/exit, mode switches, batch start/end)
        /// that an AI debugging agent needs to reconstruct the timeline of operations.
        /// </summary>
        public void LogModeTransition(string eventDescription)
        {
            ErrorHandler.LogInfo($"[PERF:MODE] {eventDescription}");
        }

        /// <summary>Start a timer using the current ErrorHandler.CallStackDepth as call level.</summary>
        public void StartTimer(string timerName)
        {
            if (!IsEnabled) return;
            if (string.IsNullOrWhiteSpace(timerName)) return;

            var level = ErrorHandler.CallStackDepth;
            var entry = new TimerInfo(timerName, level);
            entry.StartTimer();

            try
            {
                // Track running stopwatch per name (aggregate duration for this start/stop pair)
                var sw = new Stopwatch();
                if (_running.TryAdd(timerName, sw))
                {
                    sw.Start();
                }
                else
                {
                    // Already running; ignore to avoid double-start
                }

                // Append to list for this name
                var list = _timers.GetOrAdd(timerName, _ => new List<TimerInfo>(capacity: 8));
                lock (_listGate)
                {
                    list.Add(entry);
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(nameof(PerformanceTracker), "Exception in StartTimer", ex);
            }
        }

        /// <summary>Stop a timer by name. Records elapsed time on the latest entry for that name.</summary>
        public void StopTimer(string timerName)
        {
            if (!IsEnabled) return;
            if (string.IsNullOrWhiteSpace(timerName)) return;

            try
            {
                if (_running.TryRemove(timerName, out var sw))
                {
                    sw.Stop();
                }

                if (_timers.TryGetValue(timerName, out var list))
                {
                    TimerInfo last = null;
                    lock (_listGate)
                    {
                        if (list.Count > 0) last = list[list.Count - 1];
                    }
                    last?.StopTimer();
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(nameof(PerformanceTracker), "Exception in StopTimer", ex);
            }
        }

        /// <summary>
        /// Export timers to CSV sorted by CallLevel then StartTime.
        /// </summary>
        public bool ExportToCsv(string filePath)
        {
            try
            {
                if (!IsEnabled) return false;
                if (string.IsNullOrWhiteSpace(filePath)) return false;

                List<TimerInfo> snapshot;
                snapshot = _timers.Values.SelectMany(l => { lock (_listGate) { return l.ToList(); } }).ToList();
                var ordered = snapshot.OrderBy(t => t.CallLevel).ThenBy(t => t.StartTime).ToList();

                var sb = new StringBuilder();
                sb.AppendLine("Name,Level,StartLocalISO,ElapsedMs");
                foreach (var t in ordered)
                {
                    sb.AppendLine(t.ToString());
                }

                var dir = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
                File.WriteAllText(filePath, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(nameof(PerformanceTracker), "Exception in ExportToCsv", ex);
                return false;
            }
        }

        public void ClearAllTimers()
        {
            try
            {
                _running.Clear();
                lock (_listGate) { _timers.Clear(); }
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(nameof(PerformanceTracker), "Exception in ClearAllTimers", ex);
            }
        }

        public int GetTimerCount()
        {
            lock (_listGate)
            {
                return _timers.Sum(kv => kv.Value.Count);
            }
        }

        /// <summary>Returns average elapsed milliseconds for the given timer name across all occurrences.</summary>
        public double GetAverageTime(string timerName)
        {
            if (string.IsNullOrWhiteSpace(timerName)) return 0.0;
            lock (_listGate)
            {
                if (!_timers.TryGetValue(timerName, out var list) || list.Count == 0) return 0.0;
                return list.Average(t => (double)t.ElapsedMs);
            }
        }

        /// <summary>
        /// Generates a formatted timing summary showing stats for each timer.
        /// </summary>
        /// <param name="exportPath">Optional path to export CSV. If null, only returns summary string.</param>
        /// <returns>Formatted summary string with timing statistics.</returns>
        public string PrintSummary(string exportPath = null)
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== Performance Timing Summary ===");
            sb.AppendLine();

            List<TimerSummary> summaries;
            lock (_listGate)
            {
                summaries = _timers.Select(kv =>
                {
                    var entries = kv.Value.Where(t => t.ElapsedMs >= 0).ToList();
                    if (entries.Count == 0) return null;

                    return new TimerSummary
                    {
                        Name = kv.Key,
                        Count = entries.Count,
                        TotalMs = entries.Sum(t => (double)t.ElapsedMs),
                        MinMs = entries.Min(t => (double)t.ElapsedMs),
                        MaxMs = entries.Max(t => (double)t.ElapsedMs),
                        AvgMs = entries.Average(t => (double)t.ElapsedMs)
                    };
                })
                .Where(s => s != null)
                .OrderByDescending(s => s.TotalMs)
                .ToList();
            }

            if (summaries.Count == 0)
            {
                sb.AppendLine("No timing data collected.");
                return sb.ToString();
            }

            // Header
            sb.AppendLine($"{"Timer Name",-30} {"Count",6} {"Total",10} {"Avg",10} {"Min",10} {"Max",10}");
            sb.AppendLine(new string('-', 80));

            foreach (var s in summaries)
            {
                sb.AppendLine($"{TruncateName(s.Name, 30),-30} {s.Count,6} {FormatMs(s.TotalMs),10} {FormatMs(s.AvgMs),10} {FormatMs(s.MinMs),10} {FormatMs(s.MaxMs),10}");
            }

            sb.AppendLine(new string('-', 80));

            // Grand total
            double grandTotal = summaries.Sum(s => s.TotalMs);
            sb.AppendLine($"{"TOTAL",-30} {summaries.Sum(s => s.Count),6} {FormatMs(grandTotal),10}");
            sb.AppendLine();

            // Export to CSV if path provided
            if (!string.IsNullOrWhiteSpace(exportPath))
            {
                ExportToCsv(exportPath);
                sb.AppendLine($"Detailed timings exported to: {exportPath}");
            }

            return sb.ToString();
        }

        /// <summary>
        /// Logs timing summary to ErrorHandler debug output.
        /// Uses LogInfo so the summary is always captured in the log file,
        /// even when verbose debug tracing is off. This is critical for AI
        /// debugging agents that analyze performance after the run.
        /// </summary>
        public void LogSummary()
        {
            if (!IsEnabled) return;
            var summary = PrintSummary();
            foreach (var line in summary.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                ErrorHandler.LogInfo($"[PERF] {line}");
            }

            // Flag red-flag timings for AI readability
            var summaries = GetTimingSummaries();
            foreach (var s in summaries)
            {
                string flag = ClassifyTiming(s.Name, s.AvgMs);
                if (flag != null)
                {
                    ErrorHandler.LogInfo($"[PERF:FLAG] {s.Name}: avg={s.AvgMs:F0}ms - {flag}");
                }
            }
        }

        /// <summary>
        /// Checks a timer against known performance targets. Returns a warning
        /// string if the timing exceeds the red-flag threshold, null otherwise.
        /// Thresholds are from CLAUDE.md performance targets.
        /// </summary>
        private static string ClassifyTiming(string name, double avgMs)
        {
            if (name == null) return null;
            var n = name.ToLowerInvariant();

            if (n.StartsWith("insertbends2") && avgMs > 2000) return "RED FLAG: InsertBends2 target <500ms, red >2000ms";
            if (n.StartsWith("tryflatten") && avgMs > 1000) return "RED FLAG: TryFlatten target <200ms, red >1000ms";
            if (n.StartsWith("customproperty") && avgMs > 100) return "RED FLAG: CustomProperty target <50ms, red >100ms";
            if (n == "getlargestface" && avgMs > 2000) return "RED FLAG: GetLargestFace target <500ms, red >2000ms";
            if (n == "runsinglepartdata" && avgMs > 10000) return "RED FLAG: Total per-part target <3000ms, red >10000ms";

            return null;
        }

        private static string FormatMs(double ms)
        {
            if (ms < 1000)
                return $"{ms:F1}ms";
            else if (ms < 60000)
                return $"{ms / 1000.0:F2}s";
            else
                return $"{ms / 60000.0:F2}m";
        }

        private static string TruncateName(string name, int maxLen)
        {
            if (string.IsNullOrEmpty(name)) return "";
            return name.Length <= maxLen ? name : name.Substring(0, maxLen - 3) + "...";
        }

        /// <summary>
        /// Returns a list of timer summaries for external consumption (e.g., QA reports).
        /// </summary>
        public List<TimerSummary> GetTimingSummaries()
        {
            lock (_listGate)
            {
                return _timers.Select(kv =>
                {
                    var entries = kv.Value.Where(t => t.ElapsedMs >= 0).ToList();
                    if (entries.Count == 0) return null;

                    return new TimerSummary
                    {
                        Name = kv.Key,
                        Count = entries.Count,
                        TotalMs = entries.Sum(t => (double)t.ElapsedMs),
                        MinMs = entries.Min(t => (double)t.ElapsedMs),
                        MaxMs = entries.Max(t => (double)t.ElapsedMs),
                        AvgMs = entries.Average(t => (double)t.ElapsedMs)
                    };
                })
                .Where(s => s != null)
                .OrderByDescending(s => s.TotalMs)
                .ToList();
            }
        }

        /// <summary>
        /// Summary statistics for a single timer category.
        /// </summary>
        public sealed class TimerSummary
        {
            public string Name { get; set; }
            public int Count { get; set; }
            public double TotalMs { get; set; }
            public double MinMs { get; set; }
            public double MaxMs { get; set; }
            public double AvgMs { get; set; }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            ClearAllTimers();
        }
    }
}
