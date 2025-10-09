using System.Threading;

namespace NM.SwAddin.Validation
{
    /// <summary>
    /// Session-scoped validation counters with a simple summary string.
    /// Thread-safe increments; intended for lightweight logging.
    /// </summary>
    public static class ValidationStats
    {
        private static int _pass;
        private static int _fail;

        public static void Record(bool success)
        {
            if (success) Interlocked.Increment(ref _pass);
            else Interlocked.Increment(ref _fail);
        }

        public static void Clear()
        {
            Interlocked.Exchange(ref _pass, 0);
            Interlocked.Exchange(ref _fail, 0);
        }

        public static (int pass, int fail) Snapshot()
        {
            // Volatile read via Interlocked operations for safety
            int p = Interlocked.CompareExchange(ref _pass, 0, 0);
            int f = Interlocked.CompareExchange(ref _fail, 0, 0);
            return (p, f);
        }

        public static string BuildSummary()
        {
            var (p, f) = Snapshot();
            int total = p + f;
            double pct = total > 0 ? (100.0 * p / total) : 0.0;
            return $"Validation Summary: {p}/{total} passed ({pct:F1}%)";
        }
    }
}
