using System;
using System.Collections.Generic;

namespace NM.Core.DataModel
{
    /// <summary>
    /// Result of processing a single part during QA test run.
    /// All measurements in imperial units (inches, lbs) for human readability in JSON.
    /// </summary>
    public sealed class QATestResult
    {
        // Identity
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string Configuration { get; set; }

        // Processing outcome
        public string Status { get; set; }      // "Success", "Failed", "Error"
        public string Message { get; set; }
        public string Classification { get; set; }  // "SheetMetal", "Tube", "Generic", "Unknown"

        // Geometry (imperial units for readability)
        public double? Thickness_in { get; set; }
        public double? Mass_lb { get; set; }
        public double? BBoxLength_in { get; set; }
        public double? BBoxWidth_in { get; set; }
        public double? BBoxHeight_in { get; set; }

        // Sheet metal specifics
        public int? BendCount { get; set; }
        public bool? BendsBothDirections { get; set; }
        public double? FlatArea_sqin { get; set; }
        public double? CutLength_in { get; set; }

        // Tube specifics
        public double? TubeOD_in { get; set; }
        public double? TubeWall_in { get; set; }
        public double? TubeID_in { get; set; }
        public double? TubeLength_in { get; set; }
        public string TubeNPS { get; set; }
        public string TubeSchedule { get; set; }

        // Assembly specifics
        public bool? IsAssembly { get; set; }
        public int? TotalComponentCount { get; set; }
        public int? UniquePartCount { get; set; }
        public int? SubAssemblyCount { get; set; }

        // Material
        public string Material { get; set; }
        public string MaterialCategory { get; set; }

        // Costing (dollars)
        public double? MaterialCost { get; set; }
        public double? LaserCost { get; set; }
        public double? BendCost { get; set; }
        public double? TapCost { get; set; }
        public double? DeburCost { get; set; }
        public double? TotalCost { get; set; }

        // Cost breakdown by work center (for detailed comparison)
        public Dictionary<string, double> CostBreakdown { get; set; }

        // Timing
        public double ElapsedMs { get; set; }
        public DateTime ProcessedAt { get; set; } = DateTime.UtcNow;

        /// <summary>
        /// Create QATestResult from a PartData, converting units for readability.
        /// </summary>
        public static QATestResult FromPartData(PartData pd, double elapsedMs)
        {
            const double M_TO_IN = 39.3701;
            const double KG_TO_LB = 2.20462;
            const double M2_TO_SQIN = 1550.0031;

            var result = new QATestResult
            {
                FileName = System.IO.Path.GetFileName(pd.FilePath),
                FilePath = pd.FilePath,
                Configuration = pd.Configuration,
                Status = pd.Status == ProcessingStatus.Success ? "Success" :
                         pd.Status == ProcessingStatus.Failed ? "Failed" : "Error",
                Message = pd.FailureReason,
                Classification = pd.Classification.ToString(),
                ElapsedMs = elapsedMs,

                // Material
                Material = pd.Material,
                MaterialCategory = pd.MaterialCategory,
            };

            // Geometry conversions (only if non-zero)
            if (pd.Thickness_m > 0)
                result.Thickness_in = pd.Thickness_m * M_TO_IN;
            if (pd.Mass_kg > 0)
                result.Mass_lb = pd.Mass_kg * KG_TO_LB;
            if (pd.BBoxLength_m > 0)
                result.BBoxLength_in = pd.BBoxLength_m * M_TO_IN;
            if (pd.BBoxWidth_m > 0)
                result.BBoxWidth_in = pd.BBoxWidth_m * M_TO_IN;
            if (pd.BBoxHeight_m > 0)
                result.BBoxHeight_in = pd.BBoxHeight_m * M_TO_IN;

            // Sheet metal data
            if (pd.Sheet.IsSheetMetal)
            {
                result.BendCount = pd.Sheet.BendCount;
                result.BendsBothDirections = pd.Sheet.BendsBothDirections;
                if (pd.Sheet.FlatArea_m2 > 0)
                    result.FlatArea_sqin = pd.Sheet.FlatArea_m2 * M2_TO_SQIN;
                if (pd.Sheet.TotalCutLength_m > 0)
                    result.CutLength_in = pd.Sheet.TotalCutLength_m * M_TO_IN;
            }

            // Tube data
            if (pd.Tube.IsTube)
            {
                if (pd.Tube.OD_m > 0)
                    result.TubeOD_in = pd.Tube.OD_m * M_TO_IN;
                if (pd.Tube.Wall_m > 0)
                    result.TubeWall_in = pd.Tube.Wall_m * M_TO_IN;
                if (pd.Tube.ID_m > 0)
                    result.TubeID_in = pd.Tube.ID_m * M_TO_IN;
                if (pd.Tube.Length_m > 0)
                    result.TubeLength_in = pd.Tube.Length_m * M_TO_IN;
                result.TubeNPS = pd.Tube.NpsText;
                result.TubeSchedule = pd.Tube.ScheduleCode;
            }

            // Costing
            if (pd.Cost.TotalMaterialCost > 0)
                result.MaterialCost = pd.Cost.TotalMaterialCost;
            if (pd.Cost.F115_Price > 0)
                result.LaserCost = pd.Cost.F115_Price;
            if (pd.Cost.F140_Price > 0)
                result.BendCost = pd.Cost.F140_Price;
            if (pd.Cost.F220_Price > 0)
                result.TapCost = pd.Cost.F220_Price;
            if (pd.Cost.F210_Price > 0)
                result.DeburCost = pd.Cost.F210_Price;
            if (pd.Cost.TotalCost > 0)
                result.TotalCost = pd.Cost.TotalCost;

            // Cost breakdown dictionary
            result.CostBreakdown = new Dictionary<string, double>();
            if (pd.Cost.F115_Price > 0) result.CostBreakdown["F115_Laser"] = pd.Cost.F115_Price;
            if (pd.Cost.F140_Price > 0) result.CostBreakdown["F140_Bend"] = pd.Cost.F140_Price;
            if (pd.Cost.F210_Price > 0) result.CostBreakdown["F210_Debur"] = pd.Cost.F210_Price;
            if (pd.Cost.F220_Price > 0) result.CostBreakdown["F220_Tap"] = pd.Cost.F220_Price;
            if (pd.Cost.F325_Price > 0) result.CostBreakdown["F325_Roll"] = pd.Cost.F325_Price;

            return result;
        }
    }

    /// <summary>
    /// Configuration file for QA test run, read from C:\Temp\nm_qa_config.json
    /// </summary>
    public sealed class QAConfig
    {
        public string InputPath { get; set; }
        public string OutputPath { get; set; }
        public string BaselinePath { get; set; }  // Optional: path to baseline manifest.json
    }

    /// <summary>
    /// Aggregate results of a full QA test run.
    /// </summary>
    public sealed class QARunSummary
    {
        public string RunId { get; set; }
        public DateTime StartedAt { get; set; }
        public DateTime CompletedAt { get; set; }
        public int TotalFiles { get; set; }
        public int Passed { get; set; }
        public int Failed { get; set; }
        public int Errors { get; set; }
        public double TotalElapsedMs { get; set; }
        public List<QATestResult> Results { get; set; } = new List<QATestResult>();
    }
}
