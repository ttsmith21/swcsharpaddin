using System.Collections.Generic;

namespace NM.Core.Config.Sections
{
    /// <summary>
    /// Application defaults and processing parameters. Loaded from nm-config.json.
    /// </summary>
    public sealed class ProcessingDefaults
    {
        public double DefaultKFactor { get; set; } = 0.44;
        public int DefaultQuantity { get; set; } = 1;
        public int MaxRetries { get; set; } = 3;
        public double NestEfficiency { get; set; } = 1.0;
        public string DefaultSheetName { get; set; } = "Sheet1";
        public bool AutoCloseExcel { get; set; } = true;
    }

    /// <summary>
    /// Logging and debug mode settings. Loaded from nm-config.json.
    /// </summary>
    public sealed class LoggingConfig
    {
        public bool LogEnabled { get; set; } = true;
        public bool ShowWarnings { get; set; }
        public bool ProductionMode { get; set; }
        public bool DebugMode { get; set; }
        public bool PerformanceMonitoring { get; set; } = true;
        public bool SolidWorksVisible { get; set; }
    }

    /// <summary>
    /// Custom property names used on SolidWorks models.
    /// </summary>
    public sealed class CustomPropertyNames
    {
        public List<string> InitialProperties { get; set; } = new List<string>
        {
            "IsSheetMetal", "IsTube", "Thickness", "Description", "Customer",
            "CustPartNumber", "CuttingType", "Drawing", "ExportDate", "F115_Hours",
            "F115_Price", "F210_Hours", "F210_Price", "Length", "MaterialCostPerLB",
            "Model", "MPNumber", "RawWeight", "Revision", "Total_Weight"
        };
    }
}
