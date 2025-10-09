# SolidWorks Automator C# Add-in Conversion Plan

## Overview
This document tracks the detailed plan for converting the VBA macro system (SP_new.bas and related modules) to a modern, maintainable C# SolidWorks add-in. It is a living document and will be updated as features are completed and requirements evolve.

---

## 1. High-Level Architecture
- Entry Point: Add-in command launches a WinForms UI for user intent selection (part, assembly, folder).
- Core Services:
  - SolidWorksAppService (ISldWorks access)
  - Logging/ErrorService
  - TimerService
  - ExcelDataLoader
  - DocumentService (open/close/save)
  - CustomPropertiesService
- Processing Pipeline:
  - User intent → Model/Assembly/Folder detection → Validation → Problem Part Handling → Good Part Processing (Sheet Metal, Tube, Fallback) → Output/Reporting

---

## 2. Feature Epics & Tasks

Phase 1: Single-Part Foundation

### Epic 1: Add-in Bootstrap & Infrastructure
- [x] Add-in shell (ISwAddin, registration, command tab)
- [x] Logging/ErrorService
- [x] TimerService
  - [x] Implement timer data model (name, start time, running flag)
  - [ ] Support nested timer stack mirroring `clsTimerData` (push/pop validation) — Optional/Deferred
  - [x] Honor configuration toggles (production mode, enable/disable timing)
  - [x] Provide Start/Stop/Clear/ClearAll API with error handling
  - [x] Generate summary output (console/log) and CSV export
- [x] Configuration loader
- [x] Initial smoke test (command launches MessageBox)

### Epic 2: Single Part Validation Pipeline
- [x] SwModelInfo class (path, type, config, state)
  - [x] Implement Init storing file path/configuration and determining model type
  - [x] Add state machine (Unprocessed→Validated→Processing→Processed/Problem)
  - [x] Provide lazy `OpenInSolidWorks` and `Close` helpers
  - [x] Track dirty state and expose context info for logging/forms
- [x] BodyValidator.ValidateSingleSolidBody
  - [x] Detect and reject multi-body solids, surface bodies, and empty geometry
  - [x] Capture failure reasons for downstream problem-part flow
- [x] Validation result logging with pass/fail counts for single-part runs

### Epic 3: Sheet Metal Conversion Core
- [x] SheetMetalConverter
  - [x] Port InsertBends/ConvertToSheetMetal logic (config toggles TBD)
  - [x] Guard against missing thickness or unsupported edge conditions
  - [x] Options-driven bend parameters (Bend Table vs K-Factor via `ProcessingOptions`)
  - [x] Bend table path resolution/fallback (network vs local, .xls/.xlsx, drive→UNC mapping, SW install folder scan)
  - [x] Fallback `ConvertToSheetMetal` path (stub)
- [x] Simple processor two-step flow
  - [x] Probe with K-Factor → undo → re-apply with bend table if available
  - [x] When using bend table, set `UseKfactor = -1` (per API) to avoid sticking with K-Factor
  - [x] Add post-insert diagnostics logger (`[SMDBG]`)
- [x] SheetMetalPreparation utilities
  - [x] Manage flatten/unflatten sequencing for analysis modules
  - [ ] Cache critical faces and bend lines needed by downstream processors
- [x] Conversion error handling to flag models for problem-part review

### Epic 0: Custom Properties Foundation
- [x] Implement property cache bridging global/config scopes
- [x] Add ReadFromSW/WriteToSW with batch updates and type handling
- [x] Expose manufacturing property helpers (IsSheetMetal, Thickness, OptiMaterial, IsTube, etc.)
- [x] Track added/modified/deleted property states for efficient saves
- [x] Validate compatibility with legacy VBA property schema and configuration-level overrides
  - [x] Legacy rule: Default/empty config writes to Global only; non-default writes to that configuration only

### Epic 4: Tube Processing Core ⚠️ STUBBED - NEEDS DEBUGGING
- [x] PipeScheduleService (OD/wall resolver + 16"×0.5 rule) — **Service complete**
- [x] TubeMaterialCodeGenerator (A36/ALNZD → BLK/HR mapping in processor) — **Service complete**
- [x] TubeCuttingParameterService (feed and pierce tables) — **Service complete**
- [x] RoundBarValidator (+ OptiMaterial standard sizes) — **Service complete**
- [x] ExternalStart integration (optional, late-bound) — **Adapter complete**
- [x] Work-center ops (F325/F140/F210) thresholds and precision — **Service complete**
- [🔧] SimpleTubeProcessor — **STUBBED: All supporting services exist; tube feature extraction needs debugging**

**CRITICAL HANDOFF ISSUE:**
All tube processing services (PipeScheduleService, TubeMaterialCodeGenerator, TubeCuttingParameterService, RoundBarValidator, ExternalStart) are implemented and ready. The blocking issue is **tube geometry extraction from SolidWorks**:

1. **Cylinder Face Detection** - Need to iterate faces and identify cylindrical surfaces
2. **Axis Extraction** - Get cylinder axis for orientation calculations
3. **Length Calculation** - Measure distance between end caps
4. **Wall Thickness** - Extract from sheet metal feature or compare OD/ID

**CURRENT STATE:** SimpleTubeProcessor.CanProcess() returns false (stubbed). All downstream services are ready to receive tube data once extraction works.

**DEBUGGING FOCUS:** `SimpleTubeProcessor.ExtractTubeGeometry()` and related SolidWorks API calls for cylindrical body analysis.

### Epic 5: Single Part Pipeline & User Experience ✅ WORKING
- [x] Material selection dialog (WinForms)
  - [x] Material radio buttons (304L, 316L, A36, 6061, 5052, 409, 2205, 309, C22, AL6XN, ALNZD, Other)
  - [x] Bend deduction options (bend table vs K-Factor)
  - [x] Output options (Create DXF, Create Drawing, Report) and SolidWorks Visible toggle
  - [ ] Custom Properties button integration (placeholder wired, needs Epic 0)
- [x] ProcessingOptions class (replaces UserSelections)
- [x] Command shows dialog and passes `ProcessingOptions` to processor
- [x] MainRunner.RunSinglePart orchestrator (validation → tube/sheet routing → save)
- [ ] ExcelDataLoader integration — Deferred to Phase 2
- [ ] User preference inputs for conversion toggles and validation options — Deferred

**STATUS: UI → Validation → Sheet Metal conversion pipeline is working end-to-end.**

---

**Phase 2: Intelligence Layering** 🔜 NEXT PRIORITY

### Epic 6: Manufacturing Intelligence
- [x] ManufacturingCalculator (weight, time, cost)
  - [x] Port efficiency-based weight calculation (thickness correction factors, manual/auto modes)
  - [x] Implement sheet percentage calculator
  - [x] F115 (laser) end-to-end: flat-pattern cut metrics, Excel-driven speeds, OP20_S/OP20_R, F115_Price
  - [x] F140/F220/F325 calculators and property writeback
  - [x] Implement order processing costs, material markup, and pricing tolerances — Constants added
  - [x] Add volume discount tiers and costing modifiers
- [x] BendAnalyzer (bend count, tonnage, direction)
  - [x] Port bend counting from sketch analysis (OneBend/SketchBend path; FlatPattern path TBD)
  - [x] Flip detection (bends in both directions)
  - [x] Setup time calculator (1.25 min/ft) — implemented in `F140Calculator`
  - [x] Rate selection (10s…400s based on weight/length) — implemented in `F140Calculator`
- [x] MaterialValidator (thickness/material matrix)
  - [x] Implement thickness limit validation per material (SS, Al, CS)
  - [x] Bend tonnage check via Excel “Bend” sheet (col 3 thickness; SSN1=5, CS=6, AL=7; match first row ≥ thickness−0.005)

### Epic 7: Material Management System
- [x] OptiMaterial Service
  - [x] Port OptiMaterial lookup (object[,] via ExcelDataLoader) and tolerance matching (±0.008")
  - [x] Material column mapping — reused existing `LaserSpeedExcelProvider` (cols 19–22 feeds, 23–26 pierces; tabs 304L/316L/309/2205/CS/AL)
  - [x] Thickness-to-OptiMaterial code resolver
- [x] Material Properties
  - [x] Material density database (hardcoded values) — `MaterialDensityDatabase`
  - [x] Material compatibility/limits — extended `MaterialValidator` with per-alloy min/max + processing notes
  - [x] Thickness limit rules per material (SS/CS/Al) — integrated
  - [ ] Adapter to write OptiMaterial and material props to SW custom properties — Pending (vNext)

### Epic 8: Feature Analysis & Sheet Metal Utilities
- [x] Tapped Hole Analyzer
  - [x] Create wizard hole feature detector
  - [x] Port cosmetic thread analyzer (diameter < 1.0 in counted for tapping)
  - [x] Implement drill diameter adjustment note for SS (no prompts)
  - [x] Calculate F220 times (setup + run) and write properties
- [x] Face Analysis
  - [x] Port face selection logic (GetSelectedFace, GetFixedFace)
  - [x] Implement fixed face detection and largest face finder
  - [x] Extract bounding box dimensions
- [x] Bend State Management
  - [x] Implement FlattenPart/UnFlattenPart
  - [x] Port SelectFlatPattern and rebuild force logic
- [x] Geometry Extraction
  - [x] Port GetThickness from SheetMetal feature
  - [x] Implement GetMass/GetDensity wrappers (mass present; density TBD)
  - [x] Create face box extraction

### Epic 9: Costing System
- [x] Total Cost Calculator
  - [x] Port all work center rates and cost formulas
  - [x] Implement order processing costs, material markup, and pricing tolerances
  - [x] Add volume discount tiers and costing modifiers
- [x] Work Center Costing
  - [x] Calculate F115 (laser)
  - [x] Calculate F140 (brake)
  - [x] Calculate F220 (tapping)
  - [x] Calculate F325 (forming) costs
  - [ ] Support custom work centers

---

**Phase 3: Scaling to Assemblies & Batches** 📋 FUTURE

### Epic 10: Assembly Preprocessing
- [x] Unique component/config listing
- [x] Suppressed/lightweight handling
- [x] Component validation
- [x] Policy: process only the active configuration during assembly runs (all-config iteration is out-of-scope by design; explicit scenarios can be added later).

### Epic 11: Problem Part Management & UI
- [x] ProblemPartManager (retry logic)
  - [x] Maintain problem collection of `SwModelInfo`/`ModelInfo` with reasons
  - [x] Surface problem summaries to UI/logging
  - [x] Support revalidation flow for user-confirmed fixes (Confirm Fix and Revalidate All wired)
- [x] ProblemParts UI (WinForms)
  - [x] Grid view with Review, Retry marking, and CSV export
  - [x] ProblemReviewDialog with suggestions, Open in SolidWorks, and Confirm Fix
  - [ ] Batch retry wired to orchestrator (vNext)
- [x] Problem part review workflow integration with single-part orchestrator (Confirm Fix runs full single-part pipeline on PASS)

### Epic 12: Folder Processing & Progress UI
- [x] Folder traversal (recursive, file filtering)
- [x] Progress form (modeless)
- [x] Safe open/close of docs
- [x] STEP/IGES/XT import handler (imports neutral files and saves to native; externalizes virtuals)
- [x] Fully wire batch pipeline
  - [x] Import neutral files as parts/assemblies appropriately, then route per doc type
  - [x] Open each discovered document silently (visible off when configured), process using `MainRunner`/AutoWorkflow
  - [x] Policy: process active configuration only for each part/assembly
  - [x] Writeback custom properties and save/close per options
  - [x] Aggregate stats and surface summary at end (success/fail/skipped + errors)

### Epic 13: Good Part Processing Pipeline
- [x] ProcessingCoordinator (iterate GoodModels) — scaffold added
- [x] Fallback processor — `GenericPartProcessor`
- [ ] Integrate sheet metal and tube module selection logic
- [x] Coordinate custom property writeback and save/close operations

---

**Phase 4: Output, Quoting & Persistence** 📋 FUTURE

### Epic 14: Output Generation
- [ ] DrawingService (flat pattern, view rotation, etch marks)
- [ ] DXF export
- [ ] ReportGenerator (Excel BOM, part lists)

### Epic 15: Quote Processing System
- [ ] QuoteProcessor (batch, restart, resume)
- [ ] Processing checkpoint system

### Epic 16: Settings & State Persistence
- [ ] Settings serialization/deserialization
- [ ] GUI state persistence

### Epic 17: Model State & Property Synchronization
- [ ] SwModelInfo lifecycle management
  - [ ] Manage transitions between Good/Problem/Processed collections
  - [ ] Persist problem descriptions and processing timestamps
  - [ ] Ensure cached ModelDoc2 instances are opened/closed safely
- [ ] Property synchronization pipeline
  - [ ] Coordinate global vs configuration property updates
  - [ ] Propagate dirty state from properties back to model wrappers
  - [ ] Batch apply property writes post-processing with rollback on failure
- [ ] Collections integration
  - [ ] Maintain shared GoodModels/ProblemParts lists of `SwModelInfo`
  - [ ] Provide summarization for reporting and UI display

---

### Epic 18: Unified PartData DTO and Export Foundation — NEW
Purpose: decouple business logic from SolidWorks custom properties; enable centralized batch export.

Tasks:
- [x] Define strongly-typed DTO `NM.Core.DataModel.PartData` (identity, classification, material, geometry, sheet/tube data, costing, totals).
- [x] Add `NM.Core.Processing.PartDataPropertyMap` to translate DTO ⇄ property bag (legacy names preserved; string conversion centralized).
- [x] Add `NM.Core.Export.ExportManager` with CSV export and simple ERP text export.
- [ ] Orchestrator integration (incremental):
  - [ ] `MainRunner` (or equivalent pipeline entry) creates and returns `PartData` per processed part.
  - [ ] At end of a part run, persist via `CustomPropertiesService` using `PartDataPropertyMap.ToProperties()` (single batched write).
  - [ ] `ProcessorFactory`/`ProcessingCoordinator` pass a `PartData` through processors which populate fields directly (migrate off inter-stage property writes).
- [x] Batch integration:
  - [x] `FolderProcessor` aggregates `List<PartData>` from each run.
  - [x] After batch, call `ExportManager.ExportToCsv(...)` and optional `ExportToErp(...)` in the target folder.
- [ ] Unit tests (NM.Core.Tests): mapping roundtrips and exporter formatting.
- [ ] Docs: update QUICK_STATUS and HANDOFF_SUMMARY with DTO/Export usage.

Notes:
- Migration strategy: initially hydrate `PartData` from existing computed values and/or a temporary readback from properties; then thread `PartData` through services to remove property coupling.
- Units: internal meters/kg; property/export edge converts to in/lb per legacy.

---

## 3. Testing Strategy
- [ ] Unit tests for all core logic (mocked SW interfaces)
- [ ] Integration tests (self-hosted SW instance)
- [x] UI smoke tests
  - [x] 304L + Bend Table (SS table) — **PASSES**: UI shows Bend Table with `bend deduction_SS.xls`
  - [ ] A36 + Bend Table (CS table) — Ready to test
  - [ ] 6061/5052 + K-Factor (textbox value) — Ready to test
  - [ ] Already Sheet Metal (graceful) — Needs validation
  - [ ] Multi-body/empty geometry (clear failure) — Needs validation
  - [x] Problem parts UI: show list, review, export; Confirm Fix and Revalidate All
  - [x] Folder processing: imports STEP/IGES/XT, externalizes virtuals, processes, and summarizes
  - [ ] Good part pipeline: coordinator + generic processor smoke (sheet metal/tube pending)
  - [ ] Property cache: reads once, single batched write using legacy Default→Global rule
- [ ] Performance/timing checks

---

## 4. Critical Implementation Details
- [x] Constants/rates from VBA (bend rates, setup, sheet size, pierce, tab spacing)
  - [x] Port all bend rates (10s, 30s, 45s, 200s, 400s), setup rates, and work center costs
  - [x] Implement pricing modifiers (material markup, tolerance multipliers)
- [x] Excel data column mappings
  - [x] Laser2022v4.xlsx (replaces legacy NewLaser.xls)
  - [x] Thickness column = 3; Data starts at row 2 (headers in row 1)
  - [x] Feed columns by material: 19 (304/316), 20 (309/2205), 21 (A36/ALNZD), 22 (6061/5052)
  - [x] Pierce columns by material: 23 (304/316), 24 (309/2205), 25 (A36/ALNZD), 26 (6061/5052)
  - [x] Worksheet tabs: 304L, 316L, 309, 2205, CS, AL (fallback to grouped sheets)
- [x] Custom property schema (selected fields now written)
  - [x] Weight: `RawWeight`, `SheetPercent` (formats 0.####)
  - [x] Manufacturing ops: `OP20_S`, `OP20_R`, `F115_Price`, `F140_*`, `F220_*`
  - [x] Compatibility: support legacy `MaterailCostPerLB` alongside `MaterialCostPerLB`
- [x] Error recovery and retry logic
  - [x] Excel loader with retries and workbook caching; OptiMaterial sheet name "OptiMaterial" or "Material"
- [x] Performance monitoring parity
  - [x] Recreate timing enable/disable toggles tied to configuration settings
  - [x] Support nested call stack tracking for diagnostics
  - [x] Provide printable summary and CSV export compatible with logging system
- [x] OP20 and waterjet parity
  - [x] Default `OP20` to `F115 - LASER` when missing; detect `155/WATERJET` to switch behavior
  - [x] Laser pierces = loops + 2 + floor(cutLength/30); Waterjet has no pierce time; Waterjet setup = 15 + (rawWeight/sheetWeight)×30
- [ ] SolidWorks visibility control
  - [x] Default batch/auto runs with `ISldWorks::Visible = false` for speed (implemented via VisibilityScope)
  - [x] Provide status indicator while hidden (ProgressForm)
  - [ ] Allow user override via UI toggle; always restore previous visibility

---

## 5. Progress Tracking
- [x] Update this document as each feature is started/completed
- [ ] Link to code, tests, and documentation for each epic

---

## 6. Open Questions / Risks
- [🔧] **BLOCKING**: Tube geometry extraction from SolidWorks API (cylinder face detection, axis, length, wall thickness)
- [ ] External COM library dependencies (geometry export, etc.)
- [ ] Drawing template and view logic mapping
- [ ] Excel template compatibility
- [x] Bend tables: .xls vs .xlsx behavior differs by SW version; we standardize on .xls and warn on .xlsx.
- [ ] STEP assembly import nuances — Mostly handled
  - Neutral files (STEP/IGES/XT) are imported; assemblies saved as native .sldasm with virtuals externalized. Remaining: naming/path policies for externalized components.

---

## 7. Handoff Notes for Programmer

### 📚 Documentation
- **Quick Status:** `QUICK_STATUS.md` - What works right now (test scenarios)
- **Handoff Summary:** `HANDOFF_SUMMARY.md` - Detailed status and next steps
- **Tube Debugging:** `docs\\TUBE_HANDOFF.md` - Tube geometry extraction guide

### What's Working (Production-Ready)
1. **Sheet Metal Processing**: Full pipeline from UI → validation → InsertBends → save
   - Bend table resolution (network paths, UNC mapping, fallbacks)
   - Two-step processing (K-Factor probe → bend table application)
   - Material selection (304L, 316L, A36, 6061, 5052, etc.)
   - Tested with `bend deduction_SS.xls` successfully

2. **Infrastructure**: Logging, timing, error handling, configuration system all working

3. **Validation**: Single solid body detection, multi-body rejection, empty geometry handling

4. **Manufacturing (substantial)**:
   - Weight calculations per legacy (efficiency/manual) and thickness multipliers
   - Sheet percentage calculator
   - Laser F115 wired (flat-pattern cut metrics + Excel-driven speeds); properties `OP20_S`, `OP20_R`, `F115_Price`; waterjet parity
   - Press brake F140 wired; `BendAnalyzer` flip detection; properties `F140_S`, `F140_R`, `F140_S_Cost`, `F140_Price`
   - Tapped holes F220 wired; properties `F220`, `F220_S`, `F220_R`, `F220_RN`, optional `F220_Note`, `F220_Price`
   - MaterialValidator with thickness limits and Bend sheet tonnage checks
   - Material management foundations: OptiMaterial resolver (±0.008" tolerance), density DB, Excel loader support
   - Assembly preprocessing foundations: unique component/config collection, suppression/lightweight handling, imported/toolbox/virtual checks; per-part summary
   - Problem parts foundations: manager, UI list, review dialog, CSV export; add-in command to open UI
   - Total cost aggregation with order costs, material markup, difficulty, and volume discounts; writes `QuoteQty`, `MaterialCostPerLB` (both spellings), `TotalPrice`

5. **Custom Properties (Epic 0) — Completed**
   - Property cache bridging global/config scopes with change tracking
   - Batch ReadFromSW/WriteToSW; single write at end of processing
   - Strongly-typed helpers: IsSheetMetal, IsTube, Thickness, ThicknessInMeters, OptiMaterial, rbMaterialType, MaterialCategory, MaterialDensity, RawWeight, SheetPercent, MaterialCostPerLB (+ legacy `MaterailCostPerLB`), QuoteQty, Difficulty, etc.
   - Legacy write rule respected: Default/empty config writes to Global; non-default writes to that configuration

### What's Stubbed (Needs Debugging)
1. **Tube Processing** (`SimpleTubeProcessor`):
   - **All services are complete and ready**: PipeScheduleService, TubeMaterialCodeGenerator, TubeCuttingParameterService, RoundBarValidator, ExternalStart
   - **Blocking issue**: Cannot extract tube geometry from SolidWorks
   - **Needs**: Cylinder face detection, axis extraction, length measurement, wall thickness detection
   - **Current behavior**: `CanProcess()` returns false; processor skips tube parts
   - **See:** `docs\\TUBE_HANDOFF.md` for detailed debugging guide

2. **F325 Wiring**:
   - Integration into `MainRunner` and `TotalCostCalculator` pending (calculator implemented)

3. **MaterialProperty adapter integration**:
   - Excel-driven population (OptiMaterial and density) pending — property writeback now supported via CustomPropertiesService

4. **Good Part Pipeline (Epic 13)**:
   - `ProcessingCoordinator` scaffold added with stats and summary
   - `ProcessorFactory` and `GenericPartProcessor` implemented
   - Integration with sheet metal/tube processors continues (module selection)

5. **DTO and Export (Epic 18)**:
   - Initial skeleton for `PartData` DTO and `ExportManager` added
   - `FolderProcessor` now aggregates and exports CSV/ERP summary
   - Orchestrator/processor integration pending across the pipeline

### File Locations
- Add-in entry: `src/NM.SwAddin/swaddin.cs`
- Main orchestrator: `src/NM.Core/Processing/MainRunner.cs`
- Sheet metal: `src/NM.Core/Processing/SimpleSheetMetalProcessor.cs`
- Tube (stubbed): `src/NM.Core/Processing/SimpleTubeProcessor.cs`
- Tube services: `src/NM.Core/Tube/*` (all complete)
- Manufacturing core: `src/NM.Core/Manufacturing/ManufacturingCalculator.cs`, `src/NM.Core/Manufacturing/CostConstants.cs`
- Bend analysis: `src/NM.SwAddin/Manufacturing/BendAnalyzer.cs`
- Excel integration: `src/NM.SwAddin/Data/ExcelDataLoader.cs`, `src/NM.SwAddin/Manufacturing/Laser/LaserSpeedExcelProvider.cs`
- Cut metrics: `src/NM.Core/Manufacturing/CutMetrics.cs`, `src/NM.SwAddin/Manufacturing/FlatPatternAnalyzer.cs`
- Materials: `src/NM.Core/Materials/OptiMaterialService.cs`, `src/NM.Core/Materials/MaterialDensityDatabase.cs`
- Geometry: `src/NM.SwAddin/Geometry/FaceAnalyzer.cs`, `src/NM.SwAddin/Geometry/BoundingBoxExtractor.cs`
- Sheet metal: `src/NM.SwAddin/SheetMetal/BendStateManager.cs`
- Assembly: `src/NM.SwAddin/Assembly/AssemblyPreprocessor.cs`, `src/NM.SwAddin/Assembly/ComponentCollector.cs`, `src/NM.SwAddin/Assembly/ComponentValidator.cs`
- Problems: `src/NM.Core/ProblemParts/ProblemPartManager.cs`, `src/NM.SwAddin/UI/ProblemPartsForm.cs`, `src/NM.SwAddin/UI/ProblemReviewDialog.cs`
- Assemblies (quantities): `src/NM.SwAddin/Assembly/AssemblyComponentQuantifier.cs`
- Folder processing (Epic 12): `src/NM.SwAddin/Processing/FolderProcessor.cs`, `src/NM.SwAddin/UI/ProgressForm.cs`, `src/NM.SwAddin/Import/StepImportHandler.cs`
- Good part pipeline (Epic 13): `src/NM.SwAddin/Processing/ProcessingCoordinator.cs`, `src/NM.SwAddin/Processing/ProcessorFactory.cs`, `src/NM.SwAddin/Processing/GenericPartProcessor.cs`, `src/NM.SwAddin/Processing/IPartProcessor.cs`
- DTO and Export (Epic 18): `src/NM.Core/DataModel/PartData.cs`, `src/NM.Core/Processing/PartDataPropertyMap.cs`, `src/NM.Core/Export/ExportManager.cs`

### Architecture Notes
- Clean separation: `NM.SwAddin` (COM interop) vs `NM.Core` (pure C# logic)
- No DI containers, kept simple per copilot-instructions.md
- All SolidWorks API calls must be on main thread (STA COM)
- Assembly processing policy: process active configuration only (by design).

---

### Notes / Follow-ups
- UI now sets bend tables to company primary O: share:
  - `O:\\Engineering Department\\Solidworks\\Sheet Metal Bend Tables\\bend deduction_SS.xls`
  - `O:\\Engineering Department\\Solidworks\\Sheet Metal Bend Tables\\bend deduction_CS.xls`
- Resolver (`BendTableResolver`) improvements:
  - Honors explicit path; maps drive letters to UNC via registry; tries .xls↔.xlsx variants
  - Scans parent directories and SW install folders for candidates
  - Detailed decision logging added
- Processor (`SimpleSheetMetalProcessor`) changes:
  - Two-step insert flow implemented (K-Factor probe → undo → final with bend table)
  - When using bend table, passes `UseKfactor = -1` (per API); volume check ±0.5%
  - Added final sheet-metal diagnostics logger (`[SMDBG]`) (best-effort)
- **Tube processing**: All supporting services complete; geometry extraction is the only blocker.

*Last updated: February 2025 (Epic 0 complete; Epic 7 foundations + Epic 8 face/bbox/select flat + F325 calculator + Epic 10 preprocessing + Epic 11 problem parts foundations + Epic 12 folder batch wired + Epic 13 scaffold + Epic 18 DTO/Export aggregation + CSV/ERP export)*
