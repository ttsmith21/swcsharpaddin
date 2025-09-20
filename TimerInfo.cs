using System;
using System.Globalization;

namespace NM.Core
{
    /// <summary>
    /// Simple timer data container for a single timing occurrence.
    /// </summary>
    public sealed class TimerInfo
    {
        /// <summary>Name of the timed operation.</summary>
        public string TimerName { get; }
        /// <summary>UTC start time of this timer occurrence.</summary>
        public DateTime StartTime { get; private set; }
        /// <summary>Elapsed time in milliseconds for this occurrence.</summary>
        public long ElapsedMs { get; private set; }
        /// <summary>Call stack level captured when the timer started.</summary>
        public int CallLevel { get; }

        public TimerInfo(string timerName, int callLevel)
        {
            if (string.IsNullOrWhiteSpace(timerName)) throw new ArgumentException("Timer name required", nameof(timerName));
            TimerName = timerName;
            CallLevel = callLevel < 0 ? 0 : callLevel;
            StartTime = DateTime.UtcNow;
            ElapsedMs = 0L;
        }

        public void StartTimer()
        {
            StartTime = DateTime.UtcNow;
            ElapsedMs = 0L;
        }

        public void StopTimer()
        {
            var end = DateTime.UtcNow;
            var elapsed = end - StartTime;
            if (elapsed.Ticks < 0) ElapsedMs = 0L; else ElapsedMs = (long)elapsed.TotalMilliseconds;
        }

        public override string ToString()
        {
            return string.Join(",",
                TimerName?.Replace(",", " ") ?? string.Empty,
                CallLevel.ToString(CultureInfo.InvariantCulture),
                StartTime.ToLocalTime().ToString("o", CultureInfo.InvariantCulture),
                ElapsedMs.ToString(CultureInfo.InvariantCulture));
        }
    }
}
