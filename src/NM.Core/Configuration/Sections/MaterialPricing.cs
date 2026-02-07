using System.Collections.Generic;

namespace NM.Core.Config.Sections
{
    /// <summary>
    /// Material prices per pound and density data. Loaded from nm-config.json.
    /// </summary>
    public sealed class MaterialPricing
    {
        public double Stainless304_PerLb { get; set; } = 1.75;
        public double Stainless316_PerLb { get; set; } = 2.25;
        public double CarbonSteel_PerLb { get; set; } = 0.55;
        public double Aluminum6061_PerLb { get; set; } = 2.50;
        public double Aluminum5052_PerLb { get; set; } = 2.35;
        public double Galvanized_PerLb { get; set; } = 0.65;
        public double DefaultCostPerLb { get; set; } = 3.5;
    }

    /// <summary>
    /// Material density data (lb/in^3). Loaded from nm-config.json.
    /// </summary>
    public sealed class MaterialDensities
    {
        public double Stainless304_316 { get; set; } = 0.289;
        public double CarbonSteelA36 { get; set; } = 0.284;
        public double Aluminum { get; set; } = 0.098;
        public double Steel_General { get; set; } = 0.284;
    }

    /// <summary>
    /// Pricing modifiers: tolerance multipliers and markups. Loaded from nm-config.json.
    /// </summary>
    public sealed class PricingModifiers
    {
        public double MaterialMarkup { get; set; } = 1.05;
        public double TightTolerance { get; set; } = 1.15;
        public double NormalTolerance { get; set; } = 1.0;
        public double LooseTolerance { get; set; } = 0.95;
        public double OrderSetup { get; set; } = 20.0;
        public double OrderRun { get; set; } = 3.0;
    }

    /// <summary>
    /// Tensile strength values by material (psi). Loaded from nm-config.json.
    /// </summary>
    public sealed class TensileStrengthTable
    {
        public Dictionary<string, double> Values { get; set; } = new Dictionary<string, double>
        {
            { "304", 75000 },
            { "316", 75000 },
            { "309", 75000 },
            { "2205", 95000 },
            { "AL6XN", 100000 },
            { "HASTELLOY", 100000 },
            { "A36", 58000 },
            { "CS", 58000 },
            { "6061", 45000 },
            { "5052", 33000 },
            { "AL", 40000 },
        };

        public double DefaultPsi { get; set; } = 60000;
    }
}
