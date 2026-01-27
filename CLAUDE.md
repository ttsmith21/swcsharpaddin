# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a SolidWorks C# add-in that automates sheet metal processing, tube processing, and manufacturing data extraction. It's a port of an extensive VBA macro to C# for better performance, stability, and maintainability.

**Business Purpose:** Automate the path from quote → routing → ERP for sheet-metal parts and assemblies at Northern Manufacturing.

## Build & Verify Commands

```powershell
# Full build + test (ALWAYS run this after changes)
.\scripts\build-and-test.ps1

# Fast incremental build (skip clean)
.\scripts\build-and-test.ps1 -SkipClean

# Build only, no tests
.\scripts\build-and-test.ps1 -SkipClean -SkipTests

# Verify add-in registration
.\scripts\verify-registration.ps1

# Sync .cs files to csproj (preview)
.\scripts\sync-csproj.ps1 -DryRun

# Sync .cs files to csproj (apply)
.\scripts\sync-csproj.ps1
```

**CRITICAL**: This project requires Visual Studio MSBuild (not `dotnet build`) because of COM interop (`RegisterForComInterop=true`). The build scripts handle this automatically.

```bash
# Manual build from command line (if scripts unavailable)
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe" swcsharpaddin.csproj /p:Configuration=Debug

# Register the add-in (run as Administrator)
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe /codebase bin\Debug\swcsharpaddin.dll
```

**Important:** Close SolidWorks before building. The DLL is locked while SW is running.

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
│   ├── build-and-test.ps1
│   ├── sync-csproj.ps1
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

The original VBA macro is in GitHub repo `ttsmith21/Solidworks-Automator-VBA`:
- `SP.bas` - Main controller (batch processing, state management)
- `modExport.bas` - ERP data export (Import.prn generation)
- `modMaterialCost.bas` / `modMaterialUpdate.bas` - Costing calculators
- `DimensionDrawing.bas` - Automated dimensioning
- `sheetmetal1.bas` - Sheet metal validation

## Git Notes

- Current branch: `working-baseline` (known good state from commit `49c819a`)
- Broken commit: `bb89a66` - has code to port incrementally
- Run `git stash` before switching branches if you have uncommitted work

## Lessons Learned

1. **dotnet build fails for COM interop** - Use VS MSBuild via build scripts
2. **Old-style csproj needs manual file management** - Use sync-csproj.ps1
3. **Analyzers help catch issues early** - Roslynator + .NET Analyzers enabled
4. **Integration tests need SW running** - They skip gracefully otherwise
5. **COM registration is separate from build success** - DLL builds without admin
