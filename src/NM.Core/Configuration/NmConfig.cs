using NM.Core.Config.Sections;

namespace NM.Core.Config
{
    /// <summary>
    /// Root configuration POCO — deserialized from nm-config.json.
    /// All business-tunable values live here; physics constants live in UnitConversions.
    /// </summary>
    public sealed class NmConfig
    {
        public int SchemaVersion { get; set; } = 1;
        public WorkCenterRates WorkCenters { get; set; } = new WorkCenterRates();
        public ManufacturingParams Manufacturing { get; set; } = new ManufacturingParams();
        public MaterialPricing MaterialPricing { get; set; } = new MaterialPricing();
        public MaterialDensities MaterialDensities { get; set; } = new MaterialDensities();
        public PricingModifiers Pricing { get; set; } = new PricingModifiers();
        public TensileStrengthTable TensileStrengths { get; set; } = new TensileStrengthTable();
        public PathConfig Paths { get; set; } = new PathConfig();
        public ProcessingDefaults Defaults { get; set; } = new ProcessingDefaults();
        public LoggingConfig Logging { get; set; } = new LoggingConfig();
        public CustomPropertyNames CustomProperties { get; set; } = new CustomPropertyNames();
    }
}
