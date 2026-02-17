namespace NM.Core.Config.Sections
{
    /// <summary>
    /// File system paths with fallback arrays. Loaded from nm-config.json.
    /// First entry that exists on disk wins.
    /// </summary>
    public sealed class PathConfig
    {
        public string[] MaterialDataPaths { get; set; } =
        {
            @"O:\Engineering Department\Solidworks\Macros\(Semi)Autopilot\Laser2022v4.xlsx"
        };

        public BendTablePaths BendTables { get; set; } = new BendTablePaths();

        public string MaterialPropertyFilePath { get; set; } =
            @"C:\Program Files\SolidWorks Corp\SolidWorks\lang\english\sldmaterials\SolidWorks Materials.sldmat";

        public string ExtractDataAddInPath { get; set; } =
            @"C:\Program Files\SolidWorks Corp\SolidWorks\Toolbox\data collector\ExtractData.dll";

        public string ErrorLogPath { get; set; } = @"C:\SolidWorksMacroLogs\ErrorLog.txt";

        /// <summary>
        /// Path to WPS lookup CSV. First existing path wins.
        /// CSV columns: WpsNumber,Process,BaseMetal1,BaseMetal2,ThicknessMinIn,ThicknessMaxIn,JointType,Code,FillerMetal,ShieldingGas,Notes
        /// </summary>
        public string[] WpsLookupPaths { get; set; } =
        {
            @"O:\Engineering Department\Solidworks\Macros\(Semi)Autopilot\WPS_Lookup.csv",
            @"C:\SolidWorksMacroLogs\WPS_Lookup.csv"
        };
    }

    public sealed class BendTablePaths
    {
        public string[] StainlessSteel { get; set; } =
        {
            @"O:\Engineering Department\Solidworks\Bend Tables\StainlessSteel.xlsx",
            @"C:\Program Files\SolidWorks Corp\SolidWorks\lang\english\Sheet Metal Bend Tables\Stainless Steel.xls"
        };

        public string[] CarbonSteel { get; set; } =
        {
            @"O:\Engineering Department\Solidworks\Bend Tables\CarbonSteel.xlsx",
            @"C:\Program Files\SolidWorks Corp\SolidWorks\lang\english\Sheet Metal Bend Tables\Steel - Mild Steel.xls"
        };

        /// <summary>
        /// Sentinel value meaning "use K-factor mode instead of a bend table file."
        /// </summary>
        public string NoneValue { get; set; } = "-1";
    }
}
