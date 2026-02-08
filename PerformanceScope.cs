using System;

namespace NM.Core
{
    /// <summary>
    /// RAII wrapper for PerformanceTracker.StartTimer/StopTimer pairs.
    /// Guarantees StopTimer is called even on early returns or exceptions.
    ///
    /// Usage:
    ///   using (new PerformanceScope("MyOperation"))
    ///   {
    ///       // timed code
    ///   }
    /// </summary>
    public sealed class PerformanceScope : IDisposable
    {
        private readonly string _timerName;
        private bool _disposed;

        public PerformanceScope(string timerName)
        {
            _timerName = timerName;
            PerformanceTracker.Instance.StartTimer(timerName);
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            PerformanceTracker.Instance.StopTimer(_timerName);
        }
    }
}
