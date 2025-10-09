namespace NM.Core.Manufacturing
{
    public sealed class BendInfo
    {
        public int Count { get; set; }
        public double LongestBendIn { get; set; } // inches
        public double MaxRadiusIn { get; set; }    // inches
        public bool NeedsFlip { get; set; }        // bends in both directions
    }
}
