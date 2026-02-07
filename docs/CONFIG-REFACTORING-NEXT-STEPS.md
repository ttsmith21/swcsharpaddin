# Configuration Refactoring — Next Steps (Windows)

Run these steps on your Windows machine with Visual Studio and SolidWorks available.

---

## Phase 1: Build & Verify (DO FIRST)

```powershell
# 1. Pull the branch
git fetch origin claude/refactor-config-variables-Dop0K
git checkout claude/refactor-config-variables-Dop0K

# 2. Restore NuGet (Newtonsoft.Json 13.0.3 was added)
nuget restore swcsharpaddin.sln
# OR in VS: right-click solution → Restore NuGet Packages

# 3. Build
.\scripts\build-and-test.ps1 -SkipClean

# 4. If build errors, run:
.\scripts\fix-errors.ps1 -AutoFix
# Then rebuild
```

### Wire up NmConfigProvider.Initialize()

In `SwAddin.cs`, find the `ConnectToSW` method and add this line **early** (before any code that reads `Configuration.*` or `CostConstants.*`):

```csharp
NM.Core.Config.NmConfigProvider.Initialize();
```

### Verify config files deploy

After building, check that these exist:
```
bin\Debug\config\nm-config.json
bin\Debug\config\nm-tables.json
```

If missing, the `<Content>` items in `.csproj` may need adjusting. Check the build output.

### Run QA

```powershell
# Close SolidWorks first
.\scripts\build-and-test.ps1 -SkipClean
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe" src\NM.BatchRunner\NM.BatchRunner.csproj /p:Configuration=Debug /v:minimal
.\src\NM.BatchRunner\bin\Debug\NM.BatchRunner.exe --qa
```

All test results should be identical to before — the facade delegates return the same values as the old `const` fields.

---

## Phase 2: Replace Excel Dependency

Goal: Make `StaticLaserSpeedProvider`, `TubeCuttingParameterService`, and `PipeScheduleService` read from `NmConfigProvider.Tables` instead of hardcoded data. This creates a single source of truth (the JSON files) and eliminates the Excel COM dependency.

### Step 2a: StaticLaserSpeedProvider → JSON tables

**File:** `src/NM.Core/Manufacturing/Laser/StaticLaserSpeedProvider.cs`

Replace the hardcoded `_ssEntries`, `_csEntries`, `_alEntries` lists with lookups into `NmConfigProvider.Tables.LaserSpeeds`:

```csharp
using NM.Core.Config;
using NM.Core.Config.Tables;

public sealed class StaticLaserSpeedProvider : ILaserSpeedProvider
{
    private const double THICKNESS_TOLERANCE = 0.005;

    public LaserSpeed GetSpeed(double thicknessIn, string materialCode)
    {
        var tables = NmConfigProvider.Tables;
        if (tables?.LaserSpeeds == null) return default;

        var entries = GetEntriesForMaterial(tables.LaserSpeeds, materialCode);
        if (entries == null || entries.Count == 0) return default;

        double tolerance = tables.LaserSpeeds.ThicknessToleranceIn;
        double threshold = thicknessIn - tolerance;

        foreach (var e in entries)
        {
            if (e.ThicknessIn >= threshold)
                return new LaserSpeed { FeedRateIpm = e.FeedRateIpm, PierceSeconds = e.PierceSeconds };
        }

        var last = entries[entries.Count - 1];
        return new LaserSpeed { FeedRateIpm = last.FeedRateIpm, PierceSeconds = last.PierceSeconds };
    }

    private static List<LaserSpeedEntry> GetEntriesForMaterial(LaserSpeedTable table, string materialCode)
    {
        var m = (materialCode ?? string.Empty).ToUpperInvariant();
        if (m.Contains("A36") || m == "CS" || m.Contains("1018") || m.Contains("1020") || m.Contains("1045"))
            return table.CarbonSteel;
        if (m.Contains("6061") || m.Contains("5052") || m.Contains("3003") || m.Contains("5083"))
            return table.Aluminum;
        if (m == "AL" || m.StartsWith("AL-") || m.EndsWith("-AL"))
            return table.Aluminum;
        return table.StainlessSteel;
    }
}
```

**Test:** Run `/qa` — laser speed results should be identical.

### Step 2b: TubeCuttingParameterService → JSON tables

**File:** `src/NM.Core/Tubes/TubeCuttingParameterService.cs`

Replace the if/else chains with `NmTablesProvider.GetTubeCuttingParams()`:

```csharp
using NM.Core.Config;
using NM.Core.Config.Tables;

public sealed class TubeCuttingParameterService
{
    // ... CutParams class stays the same ...

    public CutParams Get(string materialCategory, double wallIn)
    {
        var (cutSpeed, pierceSec, kerfIn) = NmTablesProvider.GetTubeCuttingParams(
            NmConfigProvider.Tables, materialCategory, wallIn);

        return new CutParams
        {
            KerfIn = kerfIn,
            CutSpeedInPerMin = cutSpeed,
            PierceTimeSec = pierceSec
        };
    }
}
```

**Test:** Run `/qa` — tube cutting results should be identical.

### Step 2c: Deprecate ExcelDataLoader (optional)

Once 2a and 2b are verified, the Excel path is no longer the primary data source. You can:
1. Add `[Obsolete("Use NmConfigProvider.Tables instead")]` to `ExcelDataLoader`
2. Remove the Excel COM automation code in a future PR
3. Keep it as a fallback for now if you want belt-and-suspenders

---

## Phase 3: Cleanup & Deduplication

### Step 3a: Deduplicate M_TO_IN constants

~30 files have local `const double M_TO_IN = 39.3701` or inline `* 39.3701`. Replace with:

```csharp
using static NM.Core.Constants.UnitConversions;
// Then use: MetersToInches instead of M_TO_IN
```

**Important precision note:** Some files use `39.37007874015748` (exact inverse of 0.0254) while others use `39.3701` (VBA parity). Pick one:
- `39.3701` — matches VBA, simpler, existing QA gold standards
- `39.37007874015748` — mathematically precise

Recommendation: Keep `39.3701` in `UnitConversions.MetersToInches` for VBA parity. Add a second constant if high-precision is needed:
```csharp
public const double MetersToInchesExact = 39.37007874015748;
```

**Files to update (search for `39.3701` and `M_TO_IN`):**
```
src/NM.Core/Materials/StaticOptiMaterialService.cs (2 occurrences)
src/NM.Core/DataModel/QATestResult.cs
src/NM.Core/Export/ErpExportDataBuilder.cs
src/NM.Core/Export/ExportManager.cs
src/NM.Core/Processing/PartDataPropertyMap.cs
src/NM.Core/Processing/SimpleTubeProcessor.cs
src/NM.Core/Processing/TubeProfile.cs (5 occurrences)
src/NM.SwAddin/Manufacturing/FlatPatternAnalyzer.cs
src/NM.SwAddin/Manufacturing/BendAnalyzer.cs
src/NM.SwAddin/Manufacturing/TappedHoleAnalyzer.cs
src/NM.SwAddin/Geometry/BoundingBoxExtractor.cs
src/NM.SwAddin/Pipeline/MainRunner.cs (5 occurrences)
src/NM.SwAddin/Drawing/DrawingGenerator.cs (2 occurrences)
src/NM.SwAddin/UI/ProblemWizardForm.cs
```

### Step 3b: Deduplicate KG_TO_LB constants

Same pattern. Files:
```
src/NM.Core/DataModel/QATestResult.cs
src/NM.Core/Export/ErpExportDataBuilder.cs
src/NM.Core/Export/ExportManager.cs
src/NM.Core/Processing/PartDataPropertyMap.cs
src/NM.Core/Manufacturing/ManufacturingCalculator.cs
src/NM.Core/Manufacturing/MassValidator.cs
src/NM.Core/Manufacturing/Laser/LaserCalculator.cs
src/NM.SwAddin/Pipeline/MainRunner.cs
```

### Step 3c: Add smoke tests for config loading

**File:** `src/NM.Core.Tests/ConfigTests.cs` (new)

```csharp
using NM.Core.Config;
using Xunit;

public class ConfigTests
{
    [Fact]
    public void DefaultConfig_PassesValidation()
    {
        NmConfigProvider.ResetToDefaults();
        var msgs = ConfigValidator.Validate(NmConfigProvider.Current);
        Assert.Empty(msgs.FindAll(m => m.IsError));
    }

    [Fact]
    public void DefaultConfig_HasExpectedWorkCenterRates()
    {
        NmConfigProvider.ResetToDefaults();
        Assert.Equal(120.0, NmConfigProvider.Current.WorkCenters.F115_LaserCutting);
        Assert.Equal(80.0, NmConfigProvider.Current.WorkCenters.F140_PressBrake);
    }

    [Fact]
    public void DefaultTables_HasLaserSpeedData()
    {
        NmConfigProvider.ResetToDefaults();
        // Tables won't have data from ResetToDefaults (empty lists)
        // This test verifies the JSON loading path instead:
        // NmConfigProvider.Initialize("path/to/test/config");
    }
}
```

### Step 3d: Config editor UI (optional, future)

A WinForms dialog under `src/NM.SwAddin/UI/ConfigEditorForm.cs` that:
1. Loads `nm-config.json` into a PropertyGrid or DataGridView
2. Lets engineers edit rates/prices
3. Validates on save via `ConfigValidator`
4. Writes back to JSON

This replaces the "edit Excel spreadsheet" workflow with something simpler and git-trackable.

---

## Verification Checklist

After each step, run:
```powershell
.\scripts\build-and-test.ps1 -SkipClean
```

After Phase 2 steps, also run QA:
```powershell
.\src\NM.BatchRunner\bin\Debug\NM.BatchRunner.exe --qa
```

Results should be **identical** to baseline — all changes are refactoring (same values, different source).
