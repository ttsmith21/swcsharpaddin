# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a SolidWorks C# add-in that automates sheet metal processing, tube processing, and manufacturing data extraction. It's a port of an extensive VBA macro to C# for better performance, stability, and maintainability.

**Business Purpose:** Automate the path from quote → routing → ERP for sheet-metal parts and assemblies at Northern Manufacturing.

## Build & Verify Commands

```powershell
# Full build + test (ALWAYS run this after changes)
.\scripts\build-and-test.ps1

# Quick incremental build (~5s, for rapid iteration)
.\scripts\build-quick.ps1

# Fast incremental build with tests (skip clean)
.\scripts\build-and-test.ps1 -SkipClean

# Build only, no tests
.\scripts\build-and-test.ps1 -SkipClean -SkipTests

# Analyze build errors (CS0246, CS1061, CS0535, CS7036)
.\scripts\fix-errors.ps1

# Auto-fix missing using directives
.\scripts\fix-errors.ps1 -AutoFix

# Pre-build validation (check .csproj, duplicates, braces)
.\scripts\validate-prebuild.ps1

# Analyze and fix warnings (CS0618 deprecated calls, CS0219 unused vars)
.\scripts\fix-warnings.ps1 -DryRun   # Preview
.\scripts\fix-warnings.ps1 -AutoFix  # Apply

# Verify add-in registration
.\scripts\verify-registration.ps1

# Sync .cs files to csproj (preview)
.\scripts\sync-csproj.ps1 -DryRun

# Sync .cs files to csproj (apply)
.\scripts\sync-csproj.ps1
```

## Autonomous Fix-Compile Workflow

When fixing build errors without user intervention:

```
1. Run: .\scripts\validate-prebuild.ps1 -Fix  # Check .csproj, duplicates
2. Run: .\scripts\build-quick.ps1
3. If errors:
   a. Run: .\scripts\fix-errors.ps1 -AutoFix  # Fix CS0246, CS0103
   b. Review fix-errors.ps1 output for CS1061, CS0535, CS7036 (manual fixes)
   c. Go to step 2
4. If success with warnings:
   a. Run: .\scripts\fix-warnings.ps1 -AutoFix  # Fix CS0618 deprecated calls
   b. Re-run build to verify
5. Final: Run .\scripts\build-and-test.ps1 to verify with tests
```

**Key files created on build failure:**
- `build-errors.txt` - Full list of CS#### errors for analysis

## Headless QA Validation

Run the `/qa` skill or execute these commands to validate changes against the gold standard test suite:

```powershell
# FIRST: Verify SolidWorks is NOT running (DLL is locked while SW is open)
Get-Process -Name SLDWORKS -ErrorAction SilentlyContinue | Stop-Process -ErrorAction SilentlyContinue

# Build and run QA tests (launches SolidWorks, processes 16 test parts, closes SW)
.\scripts\build-and-test.ps1 -SkipClean
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe" src\NM.BatchRunner\NM.BatchRunner.csproj /p:Configuration=Debug /v:minimal
.\src\NM.BatchRunner\bin\Debug\NM.BatchRunner.exe --qa

# Verify DLL was updated (check timestamp)
Get-Item bin\Debug\swcsharpaddin.dll | Select-Object Name, LastWriteTime
```

**Exit codes:** 0 = all passed, 1 = failures

**Important:** If the DLL timestamp is old, SolidWorks may have been running and locked the file. Close SW and rebuild.

**Test inputs:** `tests/GoldStandard_Inputs/`
- `A*` series: Validation edge cases (empty, multi-body, no material)
- `B*` series: Sheet metal parts (native, imported, rolled cylinder)
- `C*` series: Structural shapes (tubes, angle, channel, beam)

**Autonomous debugging loop:**
```
Make change → /qa → See failures → Fix → /qa → All pass → Commit
```

**CRITICAL**: This project requires Visual Studio MSBuild (not `dotnet build`) because of COM interop (`RegisterForComInterop=true`). The build scripts handle this automatically.

```bash
# Manual build from command line (if scripts unavailable)
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe" swcsharpaddin.csproj /p:Configuration=Debug

# Register the add-in (run as Administrator)
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe /codebase bin\Debug\swcsharpaddin.dll
```

**Important:** Close SolidWorks before building. The DLL is locked while SW is running.

## Performance Instrumentation

The codebase includes a performance tracking system for identifying bottlenecks and detecting regressions.

### Using PerformanceTracker

```csharp
// Add timing to any operation
PerformanceTracker.Instance.StartTimer("MyOperation");
try {
    // ... code to measure ...
} finally {
    PerformanceTracker.Instance.StopTimer("MyOperation");
}
```

**Timer naming conventions:**
- `InsertBends2_Probe` / `InsertBends2_Final_*` - Sheet metal conversion
- `TryFlatten_Probe` / `TryFlatten_Final` - Flatten operations
- `GetLargestFace` / `FindLongestLinearEdge` - Geometry analysis
- `CustomProperty_Read` / `CustomProperty_Write` - Property I/O
- `Validation` / `Classification` / `Processing` - Pipeline phases

### BatchPerformanceScope

Use the RAII wrapper for batch operations to disable graphics updates:

```csharp
using (new BatchPerformanceScope(swApp, doc, suppressFeatureTree: true))
{
    // Batch operations run with CommandInProgress=true
    // and optionally FeatureTree disabled
}
```

### Performance Analysis Workflow

```
1. Run /qa to process test parts and generate timing.csv
2. Run /perf to analyze timing data and identify bottlenecks
3. Make optimizations
4. Run /qa again to measure improvement
5. Update baseline if performance improved: tests/timing-baseline.json
```

### Performance Targets

| Operation | Target | Red Flag |
|-----------|--------|----------|
| `InsertBends2_*` | <500ms | >2000ms |
| `TryFlatten_*` | <200ms | >1000ms |
| `CustomProperty_*` | <50ms | >100ms |
| `GetLargestFace` | <500ms | >2000ms |
| Total per-part | <3000ms | >10000ms |

### Output Files

- `tests/timing.csv` - Raw timing data from last QA run
- `tests/timing-baseline.json` - Baseline for regression detection
- `tests/Run_Latest/results.json` - Includes TimingSummary section

## Target Framework

- .NET Framework 4.8.1
- Platform: AnyCPU (runs as x64 in SolidWorks)
- COM interop enabled (RegisterForComInterop)
- SolidWorks 2022+ (references in `C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS (2)\api\redist\`)

## Architecture

### Solution Layout
```
swcsharpaddin/
├── *.cs                    # Root-level source files (legacy, main add-in)
├── src/
│   ├── NM.Core/            # Pure C# logic (no COM types in signatures)
│   │   ├── Manufacturing/  # F140, F220, F325 calculators, laser cutting
│   │   ├── Materials/      # Material density, OptiMaterial service
│   │   ├── Processing/     # SimpleTubeProcessor, PartData models
│   │   └── ProblemParts/   # Problem part tracking
│   │
│   ├── NM.SwAddin/         # COM add-in (thin glue layer)
│   │   ├── Assembly/       # Assembly traversal, component collection
│   │   ├── Geometry/       # Face analysis, bounding box extraction
│   │   ├── Manufacturing/  # Bend analyzer, flat pattern analyzer
│   │   ├── Pipeline/       # MainRunner, AutoWorkflow orchestration
│   │   ├── Processing/     # FolderProcessor, part processors
│   │   ├── Properties/     # Custom properties service
│   │   ├── SheetMetal/     # BendStateManager
│   │   ├── UI/             # Forms (ProblemParts, Progress, AutoRunSettings)
│   │   └── Validation/     # PartPreflight validation
│   │
│   ├── NM.Core.Tests/      # Unit + integration tests
│   │   ├── SmokeTests.cs
│   │   └── Integration/    # Tests requiring SolidWorks
│   │
│   └── NM.StepClassifierAddin/  # STEP file classification (separate)
│
├── scripts/                # Build automation
│   ├── build-and-test.ps1  # Full build + tests, writes build-errors.txt on failure
│   ├── build-quick.ps1     # Fast incremental build (~5s)
│   ├── fix-errors.ps1      # Analyze/fix CS0246, CS1061, CS0535, CS7036 errors
│   ├── fix-warnings.ps1    # Analyze/fix CS0618, CS0219 warnings
│   ├── validate-prebuild.ps1  # Pre-build validation (.csproj, duplicates, braces)
│   ├── sync-csproj.ps1     # Sync .cs files to csproj
│   └── verify-registration.ps1
│
└── .claude/                # Claude Code settings
```

### Core Files (Root Level - Legacy)
- `SwAddin.cs` - Main add-in entry point (49KB)
- `SolidWorksApiWrapper.cs` - API wrapper utilities (70KB)
- `SheetMetalProcessor.cs` - Sheet metal conversion (77KB)
- `SolidWorksFileOperations.cs` - File operations service
- `ErrorHandler.cs` - Logging with call stack tracking
- `PerformanceTracker.cs` - Timing system with enable/disable

### Key Services
- **MainRunner** (`src/NM.SwAddin/Pipeline/`) - Single-part orchestration
- **SimpleSheetMetalProcessor** (`src/NM.Core/Processing/`) - Sheet metal conversion (WORKING)
- **SimpleTubeProcessor** (`src/NM.Core/Processing/`) - Tube processing (geometry extraction blocked)
- **BendTableResolver** - Network path resolution with fallbacks
- **CustomPropertiesService** - Read/write SolidWorks custom properties

## Current Status

### Production-Ready (✅)
- Sheet metal processing (InsertBends with bend table)
- Validation pipeline (single body detection, multi-body rejection)
- Add-in registration and command management
- Error handling with call stack tracking
- Performance tracking with timing

### Blocked (⚠️)
- **Tube geometry extraction** - Cannot get OD/ID/length from cylinder faces
- **Custom properties write** - Read works, write not implemented

### Not Started
- Manufacturing calculations (weight, time, cost)
- Assembly batch processing
- ERP export (Import.prn generation)
- DXF/drawing output

## Critical Constraints

- **Add-in GUID is locked**: `{D5355548-9569-4381-9939-5D14252A3E47}` - do not change
- **Old-style csproj**: Files must be explicitly listed in `<Compile Include="..."/>`. Run `sync-csproj.ps1` after adding new files
- **COM registration requires admin**: Build succeeds without admin, but registration is skipped
- **Close SolidWorks before building**: The DLL is locked while SW is running

## Design Principles

- **No DI containers** - Keep it simple, small classes, early returns
- **Thin glue** - SW API calls in NM.SwAddin; algorithms in NM.Core
- **STA threading** - All SolidWorks API calls on main thread only
- **Units** - Internal = meters; convert only for display
- **COM lifetime** - Do NOT blanket-call Marshal.ReleaseComObject

## Adding New Code

| Type of Code | Location | Notes |
|--------------|----------|-------|
| Pure calculation/logic | `src/NM.Core/` | No SW types, fully testable |
| SW API wrapper | `src/NM.SwAddin/` | Thin, delegates to NM.Core |
| New calculator (F140-style) | `src/NM.Core/Manufacturing/` | Follow existing pattern |
| New part processor | `src/NM.SwAddin/Processing/` | Implement `IPartProcessor` |
| UI dialog | `src/NM.SwAddin/UI/` | WinForms |
| Unit test | `src/NM.Core.Tests/` | Name: `{Class}Tests.cs` |

**After adding ANY new .cs file**: Run `.\scripts\sync-csproj.ps1`

## VBA to C# Conversion

When converting VBA macros, follow these patterns:

### Array Returns (CRITICAL - #1 source of bugs)
```vba
' VBA
Dim vBodies As Variant
vBodies = swPart.GetBodies2(swSolidBody, True)
Set swBody = vBodies(0)
```

```csharp
// C# - MUST cast, MUST null-check
var bodiesRaw = swPart.GetBodies2((int)swBodyType_e.swSolidBody, true);
if (bodiesRaw == null) return;
var bodies = ((object[])bodiesRaw).Cast<IBody2>().ToList();
```

### Quick Reference
| VBA | C# |
|-----|-----|
| `Set obj = ...` | `var obj = ...` (remove Set) |
| `Is Nothing` | `== null` |
| `ModelDoc2` | `IModelDoc2` (add I prefix) |
| `swSolidBody` | `(int)swBodyType_e.swSolidBody` |
| `Nothing` | `null` |
| `ByRef param` | `out` or `ref` parameter |

### Conversion Workflow
1. Paste VBA in comment block for reference
2. Convert line-by-line
3. Add null checks for ALL SW API returns
4. Cast all enum params to `int`
5. Run `.\scripts\build-and-test.ps1 -SkipClean`

### After Conversion - Verify No VBA Remnants

Search converted code for these patterns (should find NONE):
- `Set ` followed by `=` (VBA assignment)
- `Is Nothing` (should be `== null`)
- `Nothing` (should be `null`)
- `ModelDoc2` without `I` prefix (should be `IModelDoc2`)
- `swSolidBody` without `(int)` cast

### SolidWorks API Gotchas

| Gotcha | Solution |
|--------|----------|
| `GetBodies2` returns `null` not empty array | Always null-check before casting |
| `CustomPropertyManager("")` | Use indexer: `CustomPropertyManager[""]` |
| Enum params need `(int)` cast | `(int)swBodyType_e.swSolidBody` |
| `GetType()` returns int, not enum | Cast: `(swDocumentTypes_e)model.GetType()` |
| Feature names may be localized | Don't hardcode "Cut-Extrude1" |
| `SelectByID2` returns bool | Check return value for success |
| Suppress state affects API calls | Check `IComponent2.IsSuppressed()` |
| Lightweight components return null | Call `ResolveAllLightweightComponents()` first |

## SolidWorks API Object Traversal

Use this map to determine the sequence of objects required to reach your target data.

### Entry Point
Everything starts with the **SldWorks** application object.
- **To get Active File:** `ISldWorks` → `IModelDoc2`
- **To get Math Tools:** `ISldWorks` → `IMathUtility`

### Document Level (IModelDoc2)
The `IModelDoc2` interface is the parent for Parts, Assemblies, and Drawings.
- **To Select Objects:** `IModelDoc2` → `ISelectionMgr`
- **To Sketch:** `IModelDoc2` → `ISketchManager`
- **To Handle Configurations:** `IModelDoc2` → `IConfigurationManager` → `IConfiguration`
- **To Edit Properties/Graphics:** `IModelDoc2` → `IModelDocExtension`

### Geometry Traversal (The "Body" Stack)
Use this flow to drill down from a file to specific topology (faces/edges).
- **From Part File:** `IPartDoc` → `IBody2` (via `GetBodies2`)
- **From Assembly Component:** `IComponent2` → `IBody2`
- **From Body:** `IBody2` → `IFace2`, `IEdge`, or `IVertex`
- **From Face:** `IFace2` → `ISurface`, `ILoop2`, or `IEdge`
- **From Edge:** `IEdge` → `ICurve` or `IVertex`

### Feature Data Editing
To edit a feature, access its specific "Definition" data object.
- **General Flow:** `IFeature` → `GetDefinition()` → *SpecificFeatureData*

| Feature Type | Data Object | Key Accessors |
|--------------|-------------|---------------|
| Extrude | `IExtrudeFeatureData2` | `SketchContour`, `SketchRegion` |
| Revolve | `IRevolveFeatureData2` | `RefAxis`, `SketchContour` |
| Sweep | `ISweepFeatureData` | `RefPlane`, `IBody2` |
| Simple Hole | `ISimpleHoleFeatureData2` | `Vertex`, `IFace2` |
| Hole Wizard | `IWizardHoleFeatureData2` | `SketchPoint`, `IFace2`, `Vertex` |
| Fillet | `ISimpleFilletFeatureData2` | `IEdge`, `IFace2`, `ILoop2` |
| Pattern | `ILinearPatternFeatureData` | `RefAxis`, `IFace2`, `IMathTransform` |

### Assembly Structure
- **To Traverse Tree:** `IAssemblyDoc` → `IComponent2`
- **To Get Sub-Components:** `IComponent2` → `IComponent2` (Children via `GetChildren()`)
- **To Get Mates:** `IAssemblyDoc` → `IMate2`
- **To Analyze Clashes:** `IAssemblyDoc` → `IInterferenceDetectionMgr`

### Drawing & Detailing
- **To get Views:** `IDrawingDoc` → `IView`
- **To get Drawing Components:** `IView` → `IDrawingComponent`
- **To get Annotations:** `IView` → `INote`, `IDisplayDimension`, `IBomTableAnnotation`

### Sketching
- **Manager:** `ISketchManager` (accessed via `IModelDoc2`)
- **Sketch Objects:** `ISketch` → `ISketchSegment`, `ISketchPoint`, `ISketchContour`

### Common Traversal Patterns

```csharp
// Get all solid bodies from a part
var partDoc = (IPartDoc)modelDoc;
var bodiesRaw = partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
if (bodiesRaw == null) return;
var bodies = ((object[])bodiesRaw).Cast<IBody2>().ToList();

// Get all faces from a body
var facesRaw = body.GetFaces();
if (facesRaw == null) return;
var faces = ((object[])facesRaw).Cast<IFace2>().ToList();

// Get surface type from face
var surface = (ISurface)face.GetSurface();
bool isCylinder = surface.IsCylinder();
bool isPlane = surface.IsPlane();

// Traverse assembly components
var config = modelDoc.ConfigurationManager.ActiveConfiguration;
var rootComp = (IComponent2)config.GetRootComponent3(true);
var children = (object[])rootComp.GetChildren();
foreach (IComponent2 child in children) { ... }
```

## SolidWorks API C# Signatures

### Application Entry Point
```csharp
using SolidWorks.Interop.sldworks;
using System.Runtime.InteropServices;

// Connect to running instance
SldWorks swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");

// Visibility Control (batch processing speedup)
swApp.Visible = true; // Set false to bypass graphics rendering
```

### ISldWorks (The App)

| Method | C# Signature | Notes |
|--------|--------------|-------|
| Open File | `ModelDoc2 OpenDoc6(string FileName, int Type, int Options, string Config, ref int Errors, ref int Warnings)` | Use `ref` for errors/warnings. Returns null if failed. |
| Close File | `void CloseDoc(string Name)` | Requires file name (e.g., "Part1.SLDPRT"), not object. |
| Activate | `int ActivateDoc3(string Name, bool Rebuild, int Opts, ref int Errors)` | Activates already-open document. |
| User Msg | `int SendMsgToUser(string Message)` | Displays dialog box. |
| Active Doc | `ModelDoc2 ActiveDoc { get; }` | Property to get current file. |

**Document Types for OpenDoc6:** `swDocPART` (1), `swDocASSEMBLY` (2), `swDocDRAWING` (3)

### IModelDoc2 (The File)
```csharp
ModelDoc2 swModel = swApp.ActiveDoc;

// Saving
bool result = swModel.Save3(int Options, ref int Errors, ref int Warnings);

// Rebuilding (TopOnly = true is faster for assemblies)
bool result = swModel.ForceRebuild3(bool TopOnly);

// Metadata
string path = swModel.GetPathName(); // Full absolute path
string title = swModel.GetTitle();   // Name in Feature Tree
int type = swModel.GetType();        // Returns swDocumentTypes_e enum

// Casting to specialized interfaces (REQUIRED for type-specific features)
PartDoc swPart = (PartDoc)swModel;
AssemblyDoc swAssy = (AssemblyDoc)swModel;
DrawingDoc swDraw = (DrawingDoc)swModel;
```

### IModelDocExtension & Selection
Modern API calls (Selection, Custom Properties) live here, not in `ModelDoc2`.
```csharp
ModelDocExtension swExt = swModel.Extension;

// SelectByID2 - The Universal Selector
// Args: Name, Type, X, Y, Z, Append, Mark, Callout, SelectOption
bool status = swModel.Extension.SelectByID2(
    "Sketch1",      // Name
    "SKETCH",       // Type
    0, 0, 0,        // X, Y, Z (0 for non-viewport selection)
    false,          // Append to current selection?
    0,              // Mark
    null,           // Callout
    0               // SelectOption
);
```

### IFeature Traversal
Do NOT rely on `Name` (e.g., "Extrude1") - it changes across languages. Use `GetTypeName2`.

| Method | Signature | Notes |
|--------|-----------|-------|
| Get First | `Feature FirstFeature()` | Via `IModelDoc2` |
| Get Next | `Feature GetNextFeature()` | Via `IFeature` (linked list) |
| Get Type | `string GetTypeName2()` | Stable ID (e.g., "ProfileFeature", "Extrusion") |
| Definition | `object GetDefinition()` | Returns data object (e.g., `ExtrudeFeatureData2`) |

```csharp
Feature swFeat = swModel.FirstFeature();
while (swFeat != null)
{
    string typeName = swFeat.GetTypeName2();
    // Logic here...
    swFeat = swFeat.GetNextFeature();
}
```

### Geometry (B-Rep)

| Method | Signature | Context |
|--------|-----------|---------|
| Get Bodies | `object[] GetBodies2(int BodyType, bool VisibleOnly)` | Returns array of bodies |
| Get Faces | `object[] GetFaces()` | Via `IBody2` |
| Get Vertices | `object[] GetVertices()` | Via `IBody2` or `IEdge` |

**Surface Math** - Get geometric data from a Face:
```csharp
// 1. Check type
ISurface surf = (ISurface)face.GetSurface();
bool isCylinder = surf.IsCylinder();
bool isPlane = surf.IsPlane();

// 2. Extract cylinder parameters
// Array: [OriginX, OriginY, OriginZ, VectorI, VectorJ, VectorK, Radius]
double[] cylParams = (double[])surf.CylinderParams;
double radius = cylParams[6];
```

### Assembly Structure (IComponent2)

| Method/Prop | Signature | Notes |
|-------------|-----------|-------|
| Get Path | `string GetPathName()` | Path to source file on disk |
| Instance Name | `string Name2 { get; }` | Returns "PartName-InstanceNum" |
| Config | `string ReferencedConfiguration { get; set; }` | Get/Set config of this instance |
| Underlying File | `ModelDoc2 GetModelDoc2()` | Returns null if lightweight/suppressed |

```csharp
// Traverse assembly tree via Configuration, NOT ModelDoc2
Configuration config = swModel.ConfigurationManager.ActiveConfiguration;
Component2 rootComp = config.GetRootComponent3(true); // true = resolve lightweight
object[] children = (object[])rootComp.GetChildren();
```

### Memory Management (Critical for Loops)
SOLIDWORKS uses unmanaged COM objects. In high-iteration loops, manual release prevents crashes:
```csharp
System.Runtime.InteropServices.Marshal.ReleaseComObject(swObject);
swObject = null;
```

## Custom Properties API (ICustomPropertyManager)

### Writing Properties - Use Add3 with OverwriteExisting
```csharp
// Add3 signature - the key is OverwriteExisting parameter
int result = customPropMgr.Add3(
    "PropertyName",       // Field name
    (int)swCustomInfoType_e.swCustomInfoText, // Type: Text, Date, Number, YesNo
    "PropertyValue",      // Value as string
    1                     // OverwriteExisting: 0=Fail if exists, 1=Overwrite
);
// Returns: 0=Success, 1=Failed (exists and Overwrite=0)
```

### Reading Properties - Use Get5 with Caching
```csharp
// Get5 avoids activating configuration (performance)
customPropMgr.Get5(
    "PropertyName",
    true,                 // UseCached = true for read without activating config
    out string valOut,
    out string resolvedValOut,
    out bool wasResolved
);
```

### Config vs File Level
- **File Level:** `ModelDoc2.Extension.CustomPropertyManager[""]`
- **Config Level:** `Configuration.CustomPropertyManager`

## Sheet Metal API

### Flat Pattern Access (Rollback Requirement)
Cannot modify flat pattern parameters unless model is in rollback state:
```csharp
// Put model in rollback state
flatPatternFeatureData.IAccessSelections2(modelDoc, null); // null for part level

// Make changes...

// MUST exit rollback state or session hangs:
feature.ModifyDefinition2(flatPatternFeatureData, modelDoc, null); // Save changes
// OR
flatPatternFeatureData.ReleaseSelectionAccess(); // Cancel changes
```

### Insert Bends
```csharp
// InsertBends2 for K-factor based conversion
bool result = partDoc.InsertBends2(
    bendRadius,           // Bend radius in meters
    bendTablePath,        // Path to .btl file, or empty for K-factor only
    kFactor,              // K-factor (use -1 if using bend table)
    bendAllowance,        // Bend allowance (use -1 if using bend table)
    useGaugeTable,        // Use gauge table
    reliefRatio,          // Relief ratio
    autoRelief            // Auto relief
);
```

### Bounding Box
Flat pattern feature must be unsuppressed. Box aligns with grain direction defined in Flat-Pattern1 feature.

## BOM Table API (IBomTableAnnotation)

### InsertBomTable3 Full Signature
```csharp
BomTableAnnotation bom = extension.InsertBomTable3(
    templatePath,         // Path to BOM template .sldbomtbt
    0,                    // X position
    0,                    // Y position
    (int)swBomType_e.swBomType_PartsOnly, // PartsOnly, TopLevel, Indented
    configName,           // CANNOT be empty for Parts/Indented types
    false,                // Hidden
    (int)swNumberingType_e.swIndentedBOMNotSet, // Numbering type
    false                 // Detailed cut list
);
```

### Correct Quantification Logic
```csharp
// 1. Get count first
int count = bomTable.GetComponentsCount2(rowIndex, configName, out string itemNo, out string modelName);

// 2. Retrieve components
object compsObj = bomTable.GetComponents2(rowIndex, configName);

// 3. Cast to IComponent2 array
if (compsObj is object[] arr && arr.Length > 0)
{
    var components = arr.Cast<IComponent2>().ToList();
    foreach (var comp in components)
    {
        string path = comp.GetPathName();
        string cfg = comp.ReferencedConfiguration;
    }
}
```

## Configuration Management

### Avoid ShowConfiguration2 in Loops
Triggers full geometry rebuild. For read-only metadata, use Get5 with UseCached=true.

### Component Config Context Trap
`IComponent2.GetModelDoc2()` returns file's last saved state, NOT the config active in assembly.
For config-specific data, access `IConfiguration` via the component:
```csharp
// Wrong - gets file default, not assembly context
var doc = component.GetModelDoc2();

// Right - get the configuration referenced by this component instance
string refConfig = component.ReferencedConfiguration;
var config = doc.GetConfigurationByName(refConfig);
var configPropMgr = config.CustomPropertyManager;
```

### Creating Configurations
```csharp
Configuration config = modelDoc.AddConfiguration3(
    "ConfigName",
    "Comment",
    "AlternateName",
    (int)swConfigurationOptions2_e.swConfigurationOption_SuppressByDefault // BOM visibility
);
```

## Code Patterns - Follow These Examples

| Task | Reference File |
|------|----------------|
| New calculator | `src/NM.Core/Manufacturing/F140Calculator.cs` |
| SW traversal | `src/NM.SwAddin/Assembly/AssemblyTraverser.cs` |
| Part processor | `src/NM.SwAddin/Processing/SheetMetalProcessor.cs` |
| Error handling | See `ErrorHandler.PushCallStack()` usage |
| Unit test | `src/NM.Core.Tests/SmokeTests.cs` |

## Success Criteria

A change is complete when:
1. `.\scripts\build-and-test.ps1` shows `BUILD: SUCCESS`
2. All tests pass (or skip gracefully if SW not running)
3. No new warnings introduced
4. New .cs files added to .csproj via `sync-csproj.ps1`
5. Code follows existing patterns in the codebase

## Debugging

```csharp
// Use ErrorHandler for call stack tracking
ErrorHandler.PushCallStack("MethodName");
try {
    // ... code ...
} finally {
    ErrorHandler.PopCallStack();
}

// Look for [SMDBG] tags in logs for sheet metal diagnostics
```

**Debug workflow:**
1. Set breakpoint in `MainRunner.ProcessActivePart()`
2. Debug → Start External Program: `SLDWORKS.exe`
3. Open test part, run "NM Classifier" → "Run Single-Part Pipeline"

## Common Type → Namespace Mappings

When you see CS0246 "type not found" errors, these are the correct namespaces:

| Type | Namespace |
|------|-----------|
| `ErrorHandler`, `ModelInfo`, `ProcessingOptions` | `NM.Core` |
| `SwModelInfo` | `NM.Core.Models` |
| `PartData`, `ProcessingStatus` | `NM.Core.DataModel` |
| `SimpleTubeProcessor`, `TubeGeometry` | `NM.Core.Processing` |
| `CutMetrics`, `TotalCostInputs` | `NM.Core.Manufacturing` |
| `ProblemPartManager` | `NM.Core.ProblemParts` |
| `PartValidationAdapter`, `PartPreflight` | `NM.SwAddin.Validation` |
| `MainRunner` | `NM.SwAddin` |
| `IModelDoc2`, `ISldWorks`, `IPartDoc` | `SolidWorks.Interop.sldworks` |
| `swDocumentTypes_e`, `swBodyType_e` | `SolidWorks.Interop.swconst` |
| `ISwAddin` | `SolidWorks.Interop.swpublished` |

The `fix-usings.ps1` script knows these mappings and can auto-add them.

## When Debugging Fails

| Issue | Solution |
|-------|----------|
| `MSB4803: RegisterAssembly not supported` | You're using `dotnet build`. Use VS MSBuild via the build script |
| `MSB3216: Cannot register assembly - access denied` | Run VS/build as Administrator, or ignore (DLL still builds) |
| `CS####` compiler errors | Check build output, fix code, rebuild |
| COM registration issues | Run `verify-registration.ps1`, then `regasm /codebase` as admin |
| DLL locked | Close SolidWorks, then rebuild |
| Tests fail with SW errors | SW must be running for integration tests |
| Missing files in build | Run `sync-csproj.ps1` to add new .cs files |

## Testing

- **Smoke tests**: Always run, test pure logic (no SW needed)
- **Integration tests**: Skip gracefully if SW not running
- Run tests: `dotnet test src/NM.Core.Tests --verbosity minimal`

## Key Documentation

- `HANDOFF_SUMMARY.md` - Current status and next steps
- `MasterPlan.md` - Full project roadmap (Epics 0-5)
- `docs/TUBE_HANDOFF.md` - Tube geometry debugging guide
- `.github/copilot-instructions.md` - Coding standards

## VBA Reference

The original VBA macro is in GitHub repo `ttsmith21/Solidworks-Automator-VBA`.
VBA source files are cached in `docs/vba-reference/` for offline access.

**See `docs/VBA-TO-CSHARP-MAPPING.md` for complete function-by-function mapping.**

| VBA Module | Purpose | Port Status |
|------------|---------|-------------|
| `SP.bas` | Main controller, batch processing | 🔶 ~60% |
| `modExport.bas` | ERP data export (Import.prn) | ❌ ~10% |
| `modMaterialCost.bas` | Cost calculations | 🔶 ~40% |
| `sheetmetal1.bas` | Sheet metal feature validation | 🔶 ~30% |
| `modConfig.bas` | Configuration constants | ✅ Done |
| `modErrorHandler.bas` | Error handling | ✅ Done |
| `FileOps.bas` | File operations | ✅ Done |

## Git Notes

- Current branch: `working-baseline`
- Code from `bb89a66` has been ported and fixed (see commits `be1d46c`, `8acd6b2`, `c71e4ec`)
- Build succeeds with 162 warnings (mostly deprecated ErrorHandler calls - harmless)
- Run `git stash` before switching branches if you have uncommitted work

## Lessons Learned

1. **dotnet build fails for COM interop** - Use VS MSBuild via build scripts
2. **Old-style csproj needs manual file management** - Use sync-csproj.ps1
3. **Analyzers help catch issues early** - Roslynator + .NET Analyzers enabled
4. **Integration tests need SW running** - They skip gracefully otherwise
5. **COM registration is separate from build success** - DLL builds without admin
