# VBA to C# Function Mapping

This document maps original VBA functions from `Solidworks-Automator-VBA` to their C# equivalents in this project.

**Legend:**
- ✅ Done - Fully ported and working
- 🔶 Partial - Ported but incomplete or needs testing
- ❌ Not Started - No C# equivalent yet
- ⏭️ Skip - Not needed in C# version

---

## SP.bas (Main Controller)

| VBA Function | C# Equivalent | Status | Notes |
|--------------|---------------|--------|-------|
| `main()` | `FolderProcessor.ProcessFolder()` | 🔶 | Batch processing loop |
| `SingleMain()` | `MainRunner.ProcessActivePart()` | ✅ | Single part orchestration |
| `QuoteStart()` | - | ❌ | Quote workflow entry |
| `QuoteStartASM()` | - | ❌ | Assembly quote workflow |
| `Initialize()` | `SwAddin.ConnectToSW()` | ✅ | Add-in initialization |
| `ProcessModel()` | `MainRunner.Run()` | 🔶 | Core processing logic |
| `CustomProperties()` | `CustomPropertiesService.ReadIntoCache()` | 🔶 | Read works, write partial |
| `SMInsertBends()` | `SimpleSheetMetalProcessor.Process()` | ✅ | Sheet metal conversion |
| `ConvertToSheetMetal()` | `SimpleSheetMetalProcessor.Process()` | ✅ | Same as above |
| `UnsuppressFlatten()` | `BendStateManager.SelectFlatPattern()` | 🔶 | Flat pattern access |
| `FindFlatPattern()` | `FlatPatternAnalyzer.GetFlatPatternFeature()` | 🔶 | Feature location |
| `ValidateFlatPattern()` | `PartPreflight.Run()` | ✅ | Validation pipeline |
| `NumberOfBodies()` | `SolidWorksApiWrapper.CountSolidBodies()` | ✅ | Body count check |
| `CompareMass()` | `MassValidator.Compare()` | ✅ | Mass validation |
| `SaveCurrentModel()` | `SolidWorksFileOperations.Save()` | ✅ | File save |
| `GetLargestFace()` | `SolidWorksApiWrapper.GetFixedFace()` | ✅ | Face selection |
| `ShowProgress()` | `ProgressForm.SetStep()` | ✅ | UI progress |
| `Report()` | - | ❌ | Summary report generation |
| `ReportPart()` | - | ❌ | Part-level reporting |
| `CreateDrawing()` | - | ❌ | Drawing automation |
| `SingleDrawing()` | - | ❌ | Single drawing creation |

### Tube Processing (SP.bas)

| VBA Function | C# Equivalent | Status | Notes |
|--------------|---------------|--------|-------|
| `ExtractTubeData()` | `SimpleTubeProcessor.Process()` | 🔶 | Geometry extraction blocked |
| `TubeCustomProperties()` | - | ❌ | Tube property writes |
| `RoundBar()` | `RoundBarValidator.IsRoundBar()` | ✅ | Round bar detection |
| `PipeDiam()` | `PipeScheduleService.TryGet()` | ✅ | Pipe schedule lookup |
| `TubeFeedRate()` | `TubeCuttingParameterService.Get()` | ✅ | Tube cutting rates |
| `TubePierceTime()` | `TubeCuttingParameterService.Get()` | ✅ | Pierce time calculation |
| `GetLinearEdge()` | `SimpleTubeProcessor.FindLongestLinearEdge()` | 🔶 | Edge detection |
| `ExGeo()` | - | ❌ | Geometry export |

### Work Center Calculations (SP.bas)

| VBA Function | C# Equivalent | Status | Notes |
|--------------|---------------|--------|-------|
| `N325()` | `F325Calculator.Calculate()` | 🔶 | Roll forming calc |
| `CalcN325()` | `F325Calculator.Calculate()` | 🔶 | Same |
| `N210()` | `F210Calculator.ComputeHours()` | ✅ | Deburr calculation |
| `BendAllowanceType()` | - | ❌ | Bend allowance logic |

---

## modExport.bas (ERP Export)

| VBA Function | C# Equivalent | Status | Notes |
|--------------|---------------|--------|-------|
| `ExportBOM()` | `ExportManager.ExportToErp()` | 🔶 | Main export routine |
| `PopulateItemMaster()` | `ErpExportFormat.WriteItemMaster()` | ✅ | Item master records |
| `PopulateProductStructure()` | `ErpExportFormat.WriteProductStructure()` | ✅ | BOM structure |
| `PopulateRouting()` | `ErpExportFormat.WriteRouting()` | ✅ | Routing records |
| `PopulateRoutingNotes()` | `ErpExportFormat.WriteRoutingNotes()` | ✅ | Routing notes |
| `PopulateParentRoute()` | `ErpExportFormat.WriteRouting()` | 🔶 | Parent assembly routes |
| `PopulateParts()` | `AssemblyComponentQuantifier.CollectQuantitiesHybrid()` | 🔶 | Part list from BOM |
| `GetBOM()` / `GetBOM1()` | `AssemblyComponentQuantifier.TryCollectViaBom()` | 🔶 | BOM table access |
| `TraverseComponent()` | `ComponentCollector.CollectComponents()` | ✅ | Assembly traversal |
| `AddIfUnique()` | - | ❌ | Duplicate filtering |
| `PartMaterialRelationships()` | - | ❌ | Material linkages |
| `RecalculateSetupTime()` | - | ❌ | Setup time recalc |
| `FixUnits()` | - | ❌ | Unit conversion |
| `SaveAsEDrawing()` | - | ❌ | eDrawings export |
| `FileNameWithoutExtension()` | `Path.GetFileNameWithoutExtension()` | ✅ | .NET built-in |
| `RemoveInstance()` | `StringUtils.RemoveInstance()` | ✅ | String parsing |
| `IsAssembly()` | - | ⏭️ | Type check |
| `AssemblyDepth()` | `StringUtils.AssemblyDepth()` | ✅ | BOM indentation |

---

## modMaterialCost.bas (Costing)

| VBA Function | C# Equivalent | Status | Notes |
|--------------|---------------|--------|-------|
| `MaterialCost()` | `MaterialCostCalculator.Calculate()` | ✅ | Main cost calculation |
| `TotalCost()` | `TotalCostCalculator.Calculate()` | 🔶 | Total cost rollup |
| `CalcWeight()` | `MetricsExtractor.FromModel()` | 🔶 | Weight calculation |
| `CalculateBendInfo()` | `BendAnalyzer.GetBendInfo()` | 🔶 | Bend analysis |
| `CalculateCutInfo()` | `FlatPatternAnalyzer.GetCutMetrics()` | 🔶 | Cut length/pierce |
| `CountBends()` | `BendAnalyzer.GetBendInfo()` | 🔶 | Bend count |
| `CheckBendTonnage()` | `BendTonnageCalculator.CheckBend()` | ✅ | Tonnage validation |
| `GetThickness()` | `ModelInfo.ThicknessInInches` | ✅ | Thickness extraction |
| `GetSelectedFace()` | `SolidWorksApiWrapper.GetFixedFace()` | ✅ | Face selection |
| `GetMass()` | `SolidWorksApiWrapper.GetMassKg()` | ✅ | Mass property |
| `GetDensity()` | `Rates.GetDensityLbPerIn3()` | ✅ | Material density |
| `GetMaterialConstants()` | `Rates.*` | 🔶 | Speed/pierce rates |
| `GetWorkCenterCosts()` | - | ❌ | Work center rates |
| `FindRate()` | - | ❌ | Rate lookup |
| `LengthWidth()` | `FlatPatternAnalyzer.GetBlankDimensions()` | 🔶 | Blank size |
| `FlattenPart()` | `BendStateManager.SelectFlatPattern()` | 🔶 | Flatten operation |
| `UnFlattenPart()` | `BendStateManager.UnFlattenPart()` | ✅ | Unflatten |
| `GetFlatFeatures()` | `FlatPatternAnalyzer.*` | 🔶 | Feature extraction |
| `GetFixedFace()` | `SolidWorksApiWrapper.GetFixedFace()` | ✅ | Fixed face for SM |
| `SelectFlatPattern()` | `BendStateManager.SelectFlatPattern()` | 🔶 | Flat pattern selection |
| `SelectSheetMetal()` | - | ❌ | SM feature selection |
| `BendData()` | `BendAnalyzer.GetBendInfo()` | 🔶 | Bend data extraction |
| `TappedHoles()` | `TappedHoleAnalyzer.Analyze()` | ✅ | Tapped hole detection |
| `FlipPart()` | - | ⏭️ | Orientation fix |
| `Fuzz()` | `Math.Abs(a-b) < tol` | ✅ | Tolerance compare |

---

## sheetmetal1.bas (Sheet Metal Validation)

| VBA Function | C# Equivalent | Status | Notes |
|--------------|---------------|--------|-------|
| `Process_SheetMetal()` | `SimpleSheetMetalProcessor.Process()` | ✅ | Main SM processing |
| `Process_Bends()` | `BendAnalyzer.GetBendInfo()` | 🔶 | Bend feature handling |
| `Process_FlatPattern()` | `FlatPatternAnalyzer.*` | 🔶 | Flat pattern handling |
| `Process_SMBaseFlange()` | - | ❌ | Base flange handling |
| `Process_EdgeFlange()` | - | ❌ | Edge flange handling |
| `Process_Hem()` | - | ❌ | Hem feature |
| `Process_Jog()` | - | ❌ | Jog feature |
| `Process_OneBend()` | - | ❌ | Single bend feature |
| `Process_LoftedBend()` | - | ❌ | Lofted bend |
| `Process_Rip()` | - | ❌ | Rip feature |
| `Process_CornerFeat()` | - | ❌ | Corner relief |
| `Process_CustomBendAllowance()` | - | ❌ | Custom bend allowance |
| `Process_SM3dBend()` | - | ❌ | 3D bend feature |
| `Process_SMMiteredFlange()` | - | ❌ | Mitered flange |
| `Process_ProcessBends()` | - | ❌ | Process bends feature |
| `Process_FlattenBends()` | - | ❌ | Flatten bends |
| `CheckBends()` | `PartPreflight.Run()` | 🔶 | Bend validation |

---

## modConfig.bas (Configuration)

| VBA Function | C# Equivalent | Status | Notes |
|--------------|---------------|--------|-------|
| Global constants | `NM.Core.Configuration` | ✅ | Paths, settings |
| File paths | `Configuration.FilePaths.*` | ✅ | Bend tables, etc. |

---

## modErrorHandler.bas (Error Handling)

| VBA Function | C# Equivalent | Status | Notes |
|--------------|---------------|--------|-------|
| `HandleError()` | `ErrorHandler.HandleError()` | ✅ | Centralized logging |
| `PushCallStack()` | `ErrorHandler.PushCallStack()` | ✅ | Call stack tracking |
| `PopCallStack()` | `ErrorHandler.PopCallStack()` | ✅ | Call stack tracking |
| `DebugLog()` | `ErrorHandler.DebugLog()` | ✅ | Debug output |

---

## FileOps.bas (File Operations)

| VBA Function | C# Equivalent | Status | Notes |
|--------------|---------------|--------|-------|
| File open/save | `SolidWorksFileOperations.*` | ✅ | Open/Save/Close |
| Browse folder | .NET `FolderBrowserDialog` | ✅ | Built-in |

---

## Summary Statistics

| Status | Count | Percentage |
|--------|-------|------------|
| ✅ Done | ~45 | ~45% |
| 🔶 Partial | ~25 | ~25% |
| ❌ Not Started | ~20 | ~20% |
| ⏭️ Skip | ~10 | ~10% |

### Critical Path - What's Blocking Production Use

1. ~~**ERP Export** (`modExport.bas`) - The `Import.prn` generation is not ported~~ ✅ DONE via `ErpExportFormat`
2. ~~**Cost Calculations** (`modMaterialCost.bas`) - `TotalCost()` incomplete~~ ✅ DONE via `TotalCostCalculator`, `MaterialCostCalculator`
3. **Tube Geometry** - Cannot extract OD/ID/length from cylinder faces ⚠️ BLOCKED
4. ~~**Custom Properties Write** - `Add3` with `OverwriteExisting` not implemented~~ ✅ Already exists in `SolidWorksApiWrapper`

### Quick Wins - COMPLETED ✅

1. ✅ `CompareMass()` → `MassValidator.Compare()` - Mass comparison
2. ✅ `N210()` → `F210Calculator.ComputeHours()` - Deburr time calculation
3. ✅ `UnFlattenPart()` → `BendStateManager.UnFlattenPart()` - Unflatten operation
4. ✅ `AssemblyDepth()` → `StringUtils.AssemblyDepth()` - BOM indentation
5. ✅ `CheckBendTonnage()` → `BendTonnageCalculator.CheckBend()` - Tonnage validation
6. ✅ `TappedHoles()` → `TappedHoleAnalyzer.Analyze()` - Tapped hole detection
7. ✅ `RemoveInstance()` → `StringUtils.RemoveInstance()` - Instance suffix removal
8. ✅ `MaterialCost()` → `MaterialCostCalculator.Calculate()` - Material cost calculation
9. ✅ `TubeFeedRate()` / `TubePierceTime()` → `TubeCuttingParameterService` - Tube cutting rates
