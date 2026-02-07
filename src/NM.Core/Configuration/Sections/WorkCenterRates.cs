namespace NM.Core.Config.Sections
{
    /// <summary>
    /// Work center hourly rates ($/hr). Loaded from nm-config.json.
    /// </summary>
    public sealed class WorkCenterRates
    {
        public double F115_LaserCutting { get; set; } = 120.0;
        public double F140_PressBrake { get; set; } = 80.0;
        public double F145_CncBending { get; set; } = 175.0;
        public double F155_Waterjet { get; set; } = 120.0;
        public double F210_Deburring { get; set; } = 42.0;
        public double F220_Tapping { get; set; } = 65.0;
        public double F300_MaterialHandling { get; set; } = 44.0;
        public double F325_RollForming { get; set; } = 65.0;
        public double F385_Assembly { get; set; } = 37.0;
        public double F400_Welding { get; set; } = 48.0;
        public double F500_Finishing { get; set; } = 48.0;
        public double F525_Packaging { get; set; } = 47.0;
        public double ENG_Engineering { get; set; } = 50.0;
    }
}
