using System;
using System.Collections.Generic;

namespace NM.Core.DataModel
{
    // Central DTO carrying all analysis/conversion/costing results for one part.
    // Internal units: meters/kg; conversions happen only at edges (property write/export).
    public sealed class PartData
    {
        // Identity
        public string FilePath { get; set; }
        public string PartName { get; set; }
        public string Configuration { get; set; }
        public string ParentAssembly { get; set; }

        // Status and classification
        public ProcessingStatus Status { get; set; } = ProcessingStatus.Pending;
        public string FailureReason { get; set; }
        public PartType Classification { get; set; } = PartType.Unknown;
        public bool IsPurchased { get; set; }

        // Material
        public string Material { get; set; }              // e.g., 304L, A36
        public string MaterialCategory { get; set; }      // SS/CS/AL
        public string OptiMaterial { get; set; }          // resolved code
        public double MaterialDensity_kg_per_m3 { get; set; }
        public double MaterialCostPerLB { get; set; }     // preserve legacy field

        // Geometry (meters/kg)
        public double Thickness_m { get; set; }
        public double Mass_kg { get; set; }
        public double SheetPercent { get; set; }          // 0..1

        // Bounding box (optional, meters)
        public double BBoxLength_m { get; set; }
        public double BBoxWidth_m { get; set; }
        public double BBoxHeight_m { get; set; }

        // Sheet metal data
        public SheetMetalData Sheet { get; } = new SheetMetalData();

        // Tube data (optional)
        public TubeData Tube { get; } = new TubeData();

        // Work centers and costing
        public CostingData Cost { get; } = new CostingData();

        // Multi-body split tracking
        /// <summary>If this part was created by splitting a multi-body part, the original part path.</summary>
        public string SplitFromParent { get; set; }
        /// <summary>Body index within the parent part (0-based). -1 if not a split part.</summary>
        public int SplitBodyIndex { get; set; } = -1;
        /// <summary>Path to the sub-assembly created from the split (set on parent PartData).</summary>
        public string SplitAssemblyPath { get; set; }

        // Commercials
        public int QuoteQty { get; set; }
        public double TotalPrice { get; set; }

        // Extensibility
        public Dictionary<string, string> Extra { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    public sealed class SheetMetalData
    {
        public bool IsSheetMetal { get; set; }
        public int BendCount { get; set; }
        public bool BendsBothDirections { get; set; }
        public double TotalCutLength_m { get; set; }
        public double FlatArea_m2 { get; set; }
    }

    public sealed class TubeData
    {
        public bool IsTube { get; set; }
        public double OD_m { get; set; }
        public double Wall_m { get; set; }
        public double ID_m { get; set; }
        public double Length_m { get; set; }

        // Additional tube properties
        public string TubeShape { get; set; } = "Round";  // Round, Square, Rectangle, Angle, Channel
        public string CrossSection { get; set; } = "";    // Cross-section description, e.g., "2.5" or "2 x 3"
        public string NpsText { get; set; }               // NPS designation, e.g., "2\"", "4\""
        public string ScheduleCode { get; set; }          // Schedule code, e.g., "40", "80S", "XX"
        public int NumberOfHoles { get; set; }            // Count of holes in the tube
        public double CutLength_m { get; set; }           // Total cut perimeter length in meters
    }

    public sealed class CostingData
    {
        // Laser / Waterjet (OP20)
        public string OP20_WorkCenter { get; set; }  // e.g., "F115", "F116"
        public double OP20_S_min { get; set; }
        public double OP20_R_min { get; set; }
        public double F115_Price { get; set; }

        // Deburr (F210)
        public double F210_S_min { get; set; }
        public double F210_R_min { get; set; }
        public double F210_Price { get; set; }

        // Press brake (F140)
        public double F140_S_min { get; set; }
        public double F140_R_min { get; set; }
        public double F140_S_Cost { get; set; }
        public double F140_Price { get; set; }

        // Tapping (F220)
        public double F220_min { get; set; }
        public double F220_S_min { get; set; }
        public double F220_R_min { get; set; }
        public int F220_RN { get; set; }
        public string F220_Note { get; set; }
        public double F220_Price { get; set; }

        // Forming (F325)
        public double F325_S_min { get; set; }
        public double F325_R_min { get; set; }
        public double F325_Price { get; set; }

        // Material costs
        public double MaterialCost { get; set; }
        public double MaterialWeight_lb { get; set; }

        // Totals
        public double TotalMaterialCost { get; set; }
        public double TotalProcessingCost { get; set; }
        public double TotalCost { get; set; }
    }

    public enum ProcessingStatus { Pending, Success, Failed, Skipped }
    public enum PartType { Unknown, SheetMetal, Tube, Generic, Assembly, Purchased }
}
