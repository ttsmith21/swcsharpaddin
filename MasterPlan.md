# Master Plan: SolidWorks Sheet Metal Automation Add-in

This document outlines the complete roadmap for porting the VBA SolidWorks automation macros to a production-ready C# add-in.

**Business Goal:** Automate the path from Quote → Routing → ERP for sheet-metal parts and assemblies at Northern Manufacturing.

---

## Ideal Workflow

The add-in supports a unified two-pass workflow that adapts to whatever is currently open in SolidWorks:

```
┌─────────────────────────────────────────────────────────────────┐
│                    USER CLICKS "RUN PIPELINE"                   │
└───────────────────────────────┬─────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────┐
│                      CONTEXT DETECTION                          │
│   What's open? → Nothing | Part | Assembly | Drawing            │
└───────────────────────────────┬─────────────────────────────────┘
                                │
        ┌───────────────────────┼───────────────────────┐
        ▼                       ▼                       ▼
   ┌─────────┐            ┌─────────┐            ┌─────────┐
   │ Nothing │            │  Part   │            │Assembly │
   │  Open   │            │  Open   │            │ or Draw │
   └────┬────┘            └────┬────┘            └────┬────┘
        │                      │                      │
        ▼                      │                      ▼
   Folder Picker               │              Collect All Parts
        │                      │              (ComponentCollector/
        ▼                      │               DrawingExtractor)
   Scan *.sldprt               │                      │
        │                      │                      │
        └──────────────────────┴──────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│                    PASS 1: VALIDATION                           │
│   BatchValidator.ValidateAll()                                  │
│   - Opens each file (if not already open)                       │
│   - Runs PartValidationAdapter.Validate()                       │
│   - Checks: single body, material, thickness, not suppressed    │
│   - Categorizes: GoodModels vs ProblemModels                    │
│   - Tracks problems in ProblemPartManager                       │
└───────────────────────────────┬─────────────────────────────────┘
                                │
                    ┌───────────┴───────────┐
                    │   Problems Found?     │
                    └───────────┬───────────┘
                          Yes   │   No
                    ┌───────────┴───────────┐
                    ▼                       │
┌─────────────────────────────┐             │
│     PROBLEM PARTS FORM      │             │
│  - Review individual parts  │             │
│  - Retry after fixes        │             │
│  - Export CSV               │             │
│  - [Continue with Good]     │◄────────────┘
│  - [Cancel]                 │
└─────────────┬───────────────┘
              │
              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    PASS 2: PROCESSING                           │
│   MainRunner.RunSinglePart() for each good model                │
│   - Detect part type (SheetMetal, Tube, Generic)                │
│   - Convert to sheet metal (InsertBends with bend table)        │
│   - Extract geometry (bounding box, flat pattern)               │
│   - Calculate costs (laser, bend, tap, deburr)                  │
│   - Write custom properties                                     │
│   - Save document                                               │
└───────────────────────────────┬─────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────┐
│                      FINAL SUMMARY                              │
│   - Total discovered, validated, processed                      │
│   - Problems encountered                                        │
│   - Timing statistics                                           │
└─────────────────────────────────────────────────────────────────┘
```

---

## Epics Overview

| Epic | Name | Description | Status |
|------|------|-------------|--------|
| 0 | Foundation | Add-in shell, error handling, file ops | ✅ Complete |
| 1 | Sheet Metal Processing | InsertBends, flat pattern, validation | ✅ Complete |
| 2 | Cost Calculations | Laser, bend, tap, deburr, material | ✅ Complete |
| 3 | ERP Export | Import.prn generation, routing records | ✅ Complete |
| 4 | Unified Workflow | Two-pass validation, context detection | ✅ Complete |
| 5 | Tube Processing | OD/ID extraction, tube-specific routing | 🔶 Partial |
| 6 | Assembly Processing | BOM traversal, quantity rollup | 🔶 Partial |
| 7 | Drawing Output | Auto-generate drawings, DXF export | 🔶 Partial |

---

## Epic 0: Foundation ✅

**Goal:** Establish the add-in infrastructure and core services.

### Completed
- [x] Add-in registration and toolbar commands
- [x] `ErrorHandler` with call stack tracking
- [x] `PerformanceTracker` for timing
- [x] `SolidWorksFileOperations` for open/save/close
- [x] `SolidWorksApiWrapper` for common API patterns
- [x] `ProcessingOptions` configuration
- [x] `ModelInfo` / `SwModelInfo` state tracking
- [x] `CustomPropertyData` cache with change tracking

### Key Files
- `SwAddin.cs` - Main entry point
- `ErrorHandler.cs` - Logging infrastructure
- `SolidWorksFileOperations.cs` - File operations
- `src/NM.Core/Models/SwModelInfo.cs` - Model state machine

---

## Epic 1: Sheet Metal Processing ✅

**Goal:** Convert imported parts to sheet metal and extract geometry.

### Completed
- [x] `SimpleSheetMetalProcessor.Process()` - InsertBends with bend table
- [x] `BendTableResolver` - Network path resolution with fallbacks
- [x] `FlatPatternAnalyzer` - Flat pattern feature access
- [x] `BendStateManager` - Flatten/unflatten operations
- [x] `BoundingBoxExtractor` - Bounding box dimensions
- [x] `PartPreflight` / `PartValidationAdapter` - Validation pipeline

### Validation Checks
1. Single solid body (multi-body rejected)
2. Material assigned
3. Thickness extractable
4. Not suppressed or lightweight
5. Not an imported/foreign file without conversion

### Key Files
- `SimpleSheetMetalProcessor.cs`
- `src/NM.SwAddin/Validation/PartPreflight.cs`
- `src/NM.SwAddin/Manufacturing/FlatPatternAnalyzer.cs`
- `src/NM.SwAddin/SheetMetal/BendStateManager.cs`

---

## Epic 2: Cost Calculations ✅

**Goal:** Calculate manufacturing costs for laser, bend, tap, and deburr operations.

### Completed
- [x] `LaserCalculator` - Laser cutting time from cut length + pierce count
- [x] `LaserSpeedService` / `LaserSpeedExcelProvider` - Material/thickness lookup
- [x] `F140Calculator` - Press brake setup/run time
- [x] `F210Calculator` - Deburr time calculation
- [x] `F220Calculator` - Tapping time calculation
- [x] `F325Calculator` - Roll forming calculation
- [x] `BendTonnageCalculator` - Tonnage validation
- [x] `MaterialCostCalculator` - Material cost from weight + rate
- [x] `TappedHoleAnalyzer` - Detect and count tapped holes
- [x] `BendAnalyzer` - Bend count and dimensions
- [x] `MetricsExtractor` - Extract all metrics from model

### Cost Data Model (`PartData.CostingData`)
```csharp
public double OP20_S_min;      // Laser setup (minutes)
public double OP20_R_min;      // Laser run (minutes)
public double F140_S_min;      // Bend setup
public double F140_R_min;      // Bend run
public double F210_S_min;      // Deburr setup
public double F210_R_min;      // Deburr run
public double F220_S_min;      // Tap setup
public double F220_R_min;      // Tap run
public double F325_S_min;      // Roll form setup
public double F325_R_min;      // Roll form run
public double F115_Price;      // Total laser cost
public double MaterialCost;    // Material cost
```

### Key Files
- `src/NM.Core/Manufacturing/LaserCalculator.cs`
- `src/NM.Core/Manufacturing/F140Calculator.cs`
- `src/NM.Core/Manufacturing/F220Calculator.cs`
- `src/NM.SwAddin/Manufacturing/MetricsExtractor.cs`

---

## Epic 3: ERP Export ✅

**Goal:** Generate Import.prn files for ERP system import.

### Completed
- [x] `ErpExportFormat` - Write PRN file with proper formatting
- [x] `ErpExportDataBuilder` - Build export data from PartData collection
- [x] `ExportManager` - CSV export for parts
- [x] `QuoteWorkflow` - Batch quote processing with ERP export option

### ERP Record Types
1. **Item Master** - Part number, description, material
2. **Product Structure** - BOM relationships (parent/child)
3. **Routing** - Work center operations (OP20, F140, F210, F220, F325)
4. **Routing Notes** - Operation notes (tap sizes, etc.)

### Key Files
- `src/NM.Core/Export/ErpExportFormat.cs`
- `src/NM.Core/Export/ErpExportDataBuilder.cs`
- `src/NM.SwAddin/Pipeline/QuoteWorkflow.cs`

---

## Epic 4: Unified Workflow ✅

**Goal:** Two-pass validation architecture with context-aware entry point.

### Completed
- [x] `WorkflowDispatcher` - Central orchestrator
- [x] `WorkflowContext` - State between passes
- [x] `BatchValidator` - Pass 1 validation
- [x] `DrawingReferenceExtractor` - Extract models from drawings
- [x] `ProblemPartsForm` - UI with "Continue with Good" button
- [x] `ProblemPartManager` - Singleton problem tracking
- [x] "Run Pipeline" toolbar command

### Context Detection
| What's Open | Action |
|-------------|--------|
| Nothing | Show folder picker, scan for *.sldprt |
| Part | Validate and process single part |
| Assembly | Collect all components, validate all, process good |
| Drawing | Extract referenced models, validate all, process good |

### Key Files
- `src/NM.SwAddin/Pipeline/WorkflowDispatcher.cs`
- `src/NM.SwAddin/Pipeline/WorkflowContext.cs`
- `src/NM.SwAddin/Validation/BatchValidator.cs`
- `src/NM.SwAddin/Drawing/DrawingReferenceExtractor.cs`

---

## Epic 5: Tube Processing 🔶

**Goal:** Process tube/pipe parts with specialized geometry extraction.

### Completed
- [x] `SimpleTubeProcessor` - Basic tube processing
- [x] `TubeGeometryExtractor` - OD/ID/length extraction
- [x] `PipeScheduleService` - Pipe schedule lookup
- [x] `TubeCuttingParameterService` - Cutting rates
- [x] `TubeMaterialCodeGenerator` - Material code generation
- [x] `RoundBarValidator` - Detect round bar vs tube

### Remaining
- [ ] Tube work center routing (different from sheet metal)
- [ ] Tube-specific custom properties
- [ ] Tube cost calculations integration

### Key Files
- `src/NM.Core/Processing/SimpleTubeProcessor.cs`
- `src/NM.SwAddin/Geometry/TubeGeometryExtractor.cs`
- `src/NM.Core/Tubes/PipeScheduleService.cs`

---

## Epic 6: Assembly Processing 🔶

**Goal:** Full assembly BOM extraction and quantity rollup.

### Completed
- [x] `ComponentCollector` - Collect unique components
- [x] `ComponentValidator` - Validate component state
- [x] `AssemblyComponentQuantifier` - Quantity extraction
- [x] `AssemblyPreprocessor` - Resolve lightweight components

### Remaining
- [ ] Full BOM table integration
- [ ] Multi-level assembly support
- [ ] Assembly-level custom properties
- [ ] Parent routing generation

### Key Files
- `src/NM.SwAddin/Assembly/ComponentCollector.cs`
- `src/NM.SwAddin/Assembly/AssemblyComponentQuantifier.cs`

---

## Epic 7: Drawing Output 🔶

**Goal:** Auto-generate drawings and DXF export.

### Completed
- [x] `DrawingGenerator` - Create drawing from part
- [x] `EDrawingExporter` - Export to eDrawings format
- [x] `ReportService` - Generate summary reports

### Remaining
- [ ] DXF export from flat pattern
- [ ] Drawing template customization
- [ ] Auto-dimensioning
- [ ] Drawing annotation

### Key Files
- `src/NM.SwAddin/Drawing/DrawingGenerator.cs`
- `src/NM.SwAddin/Export/EDrawingExporter.cs`

---

## Architecture

### Layer Separation

```
┌─────────────────────────────────────────────────────────────────┐
│                        NM.SwAddin                               │
│   Thin glue layer - SolidWorks API calls only                   │
│   - Assembly/, Drawing/, Geometry/, Manufacturing/              │
│   - Pipeline/, Processing/, Properties/, UI/, Validation/       │
└───────────────────────────────┬─────────────────────────────────┘
                                │ calls
                                ▼
┌─────────────────────────────────────────────────────────────────┐
│                         NM.Core                                 │
│   Pure C# logic - no COM types in public signatures             │
│   - DataModel/, Export/, Manufacturing/, Materials/             │
│   - Models/, ProblemParts/, Processing/, Tubes/, Utils/         │
└─────────────────────────────────────────────────────────────────┘
```

### Key Design Principles

1. **No DI containers** - Keep it simple, small classes, early returns
2. **Thin glue** - SW API calls in NM.SwAddin; algorithms in NM.Core
3. **STA threading** - All SolidWorks API calls on main thread only
4. **Units** - Internal = meters; convert only for display
5. **COM lifetime** - Do NOT blanket-call Marshal.ReleaseComObject

---

## Testing Strategy

### Unit Tests (NM.Core.Tests)
- Pure logic tests, no SolidWorks dependency
- Run with `dotnet test src/NM.Core.Tests`
- Currently: 66 tests passing

### Integration Tests
- Require SolidWorks running
- Skip gracefully if SW not available
- Test actual file operations and API calls

### Manual Testing
1. Open test part → Run Pipeline → Verify validation
2. Open assembly → Run Pipeline → Verify component collection
3. No doc open → Run Pipeline → Verify folder picker
4. Check ProblemPartsForm with known bad files

---

## Build & Deploy

```powershell
# Full build + test
.\scripts\build-and-test.ps1

# Quick incremental build
.\scripts\build-quick.ps1

# Sync new files to csproj
.\scripts\sync-csproj.ps1

# Register add-in (admin required)
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe /codebase bin\Debug\swcsharpaddin.dll
```

---

## Current Status

**Build:** ✅ SUCCESS (66/66 tests pass)

**Production Readiness:**
- Sheet metal processing: Ready
- Cost calculations: Ready
- ERP export: Ready
- Unified workflow: Ready
- Tube processing: Needs testing
- Assembly processing: Partial
- Drawing output: Partial

---

## Next Steps

1. **Test unified workflow** with real production files
2. **Complete tube routing** - work center assignments
3. **Full BOM integration** - multi-level assemblies
4. **DXF export** - flat pattern output for nesting
5. **User documentation** - training materials
