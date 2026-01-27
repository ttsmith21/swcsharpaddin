using System;

namespace NM.Core
{
    /// <summary>
    /// Immutable-like data container describing user-selected processing options for a run.
    /// Pure NM.Core model (no SolidWorks types).
    /// </summary>
    public sealed class ProcessingOptions
    {
        /// <summary>Entry point of the workflow.</summary>
        public ProcessingEntryPoint EntryPoint { get; set; }
        /// <summary>Folder path to process when EntryPoint is FolderQuote.</summary>
        public string FolderPath { get; set; }

        // Material properties
        /// <summary>SolidWorks material name (e.g., "AISI 304").</summary>
        public string MaterialType { get; set; }
        /// <summary>Internal material code (e.g., "304L").</summary>
        public string Material { get; set; }
        /// <summary>Material category (StainlessSteel, CarbonSteel, Aluminum, Other).</summary>
        public MaterialCategoryKind MaterialCategory { get; set; }
        /// <summary>Path to bend table or "-1" to use K-factor.</summary>
        public string BendTable { get; set; }
        /// <summary>User-specified K-factor; set to -1 to use bend table.</summary>
        public double KFactor { get; set; }

        // Drawing options
        /// <summary>Create DXF output during processing.</summary>
        public bool CreateDXF { get; set; }
        /// <summary>Whether to show SolidWorks UI during processing.</summary>
        public bool SolidworksVisible { get; set; }
        /// <summary>Create a drawing file.</summary>
        public bool CreateDrawing { get; set; }
        /// <summary>Generate a processing report.</summary>
        public bool GenerateReport { get; set; }

        // Custom properties
        /// <summary>Customer name for custom properties.</summary>
        public string Customer { get; set; }
        /// <summary>Print identifier for custom properties.</summary>
        public string Print { get; set; }
        /// <summary>Revision identifier for custom properties.</summary>
        public string Revision { get; set; }
        /// <summary>Description for custom properties.</summary>
        public string Description { get; set; }
        /// <summary>Use part number when generating outputs.</summary>
        public bool UsePartNum { get; set; }
        /// <summary>Honor grain direction constraints for sheet metal.</summary>
        public bool GrainConstraint { get; set; }
        /// <summary>Enable common cut (nesting) options if applicable.</summary>
        public bool CommonCut { get; set; }
        /// <summary>Save changes to the model after processing.</summary>
        public bool SaveChanges { get; set; }

        // Costing options
        /// <summary>Enable quoting/costing calculations.</summary>
        public bool QuoteEnabled { get; set; }
        /// <summary>Material cost per pound.</summary>
        public double CostPerLB { get; set; }
        /// <summary>Requested quantity for costing.</summary>
        public int Quantity { get; set; }
        /// <summary>Difficulty factor (Tight, Normal, Loose) impacting costing multipliers.</summary>
        public DifficultyLevel Difficulty { get; set; }

        /// <summary>
        /// Create a new options instance with defaults derived from Configuration constants
        /// and historical VBA defaults.
        /// </summary>
        public ProcessingOptions()
        {
            // Workflow
            EntryPoint = ProcessingEntryPoint.SinglePart;
            FolderPath = string.Empty;

            // Material defaults (from VBA Init)
            MaterialType = "AISI 304";
            Material = "304L";
            MaterialCategory = MaterialCategoryKind.StainlessSteel;
            BendTable = Configuration.FilePaths.BendTableSs; // BEND_TABLE_SS
            KFactor = -1.0; // -1 => use bend table

            // Drawing options
            CreateDXF = false;
            SolidworksVisible = false;
            CreateDrawing = false;
            GenerateReport = false;

            // Custom props
            Customer = string.Empty;
            Print = string.Empty;
            Revision = string.Empty;
            Description = string.Empty;
            UsePartNum = false;
            GrainConstraint = false;
            CommonCut = false;
            SaveChanges = true;

            // Costing
            QuoteEnabled = false;
            CostPerLB = Configuration.Defaults.DefaultCostPerLb; // DEFAULT_COST_PER_LB
            Quantity = Configuration.Defaults.DefaultQuantity;   // DEFAULT_QUANTITY
            Difficulty = DifficultyLevel.Normal;
        }
    }

    /// <summary>Entry point kinds for processing workflow.</summary>
    public enum ProcessingEntryPoint
    {
        SinglePart = 0,
        Assembly = 1,
        FolderQuote = 2
    }

    /// <summary>High-level material categories.</summary>
    public enum MaterialCategoryKind
    {
        StainlessSteel = 0,
        CarbonSteel = 1,
        Aluminum = 2,
        Other = 3
    }

    /// <summary>Difficulty factor for costing/timing multipliers.</summary>
    public enum DifficultyLevel
    {
        Tight = 0,
        Normal = 1,
        Loose = 2
    }
}
