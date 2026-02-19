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
        /// Path to drawing template (.drwdot) for auto-generated drawings.
        /// Default uses the Northern A-SIZE template on the O: drive.
        /// </summary>
        public string DrawingTemplatePath { get; set; } =
            @"O:\Engineering Department\Solidworks\Document Templates\Northern-Rev4\A-SIZE.drwdot";
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
