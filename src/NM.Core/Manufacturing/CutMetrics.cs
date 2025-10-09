namespace NM.Core.Manufacturing
{
    // Pure data object for cut metrics from a flat pattern
    public sealed class CutMetrics
    {
        public double PerimeterLengthIn { get; set; }
        public double InternalCutLengthIn { get; set; }
        public double TotalCutLengthIn => PerimeterLengthIn + InternalCutLengthIn;
        public int PierceCount { get; set; } // total loops (outer+inner) used by legacy calc
        public int HoleCount { get; set; }   // inner loops only
    }
}
