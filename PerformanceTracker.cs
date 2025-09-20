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

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            ClearAllTimers();
        }
    }
}
