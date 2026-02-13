# Plan: Fix Side Indicator Color Toggle + Automated Validation

## Problem Analysis

The `SideIndicatorService` (in `src/NM.SwAddin/SheetMetal/SideIndicatorService.cs`) has two bugs:

### Bug 1: Toggle off doesn't restore original colors
**Root cause:** `ClearBodyColors()` (line 307-329) calls `body.RemoveMaterialProperty(swAllConfiguration, null)` which **removes the body-level color** in addition to face-level overrides. If the part had a body-level or part-level color applied, this strips it entirely, leaving faces in the default system color instead of the original.

### Bug 2: No save/restore of original face colors
**Root cause:** The code never reads the original `face.MaterialPropertyValues` before overwriting with green/red/gray. If any faces had explicit face-level color overrides applied by the user, those are permanently lost. The "remove" path sets `face.MaterialPropertyValues = null` which removes the face override entirely rather than restoring the original value.

### SolidWorks Color Hierarchy (Key Insight)
```
Part color < Body color < Feature color < Face color
```
- Setting `face.MaterialPropertyValues = green` overrides all levels below
- Setting `face.MaterialPropertyValues = null` removes the face override, so body/part color shows through
- Calling `body.RemoveMaterialProperty()` removes body color AND all face overrides — destructive!

**The correct approach:** Save each face's `MaterialPropertyValues` before applying overrides. On restore, write back the saved value (`null` means "no face override was present" — set null to remove our override and let the body/part color show through).

---

## Implementation Plan

### Step 1: Add face color save/restore to SideIndicatorService

**File:** `src/NM.SwAddin/SheetMetal/SideIndicatorService.cs`

**Changes:**
1. Add a dictionary field to store original face colors per model:
   ```csharp
   // Maps model path -> list of (IFace2 reference, original double[] or null)
   private readonly Dictionary<string, List<SavedFaceColor>> _savedColors
       = new Dictionary<string, List<SavedFaceColor>>(StringComparer.OrdinalIgnoreCase);

   private class SavedFaceColor
   {
       public IFace2 Face;
       public double[] OriginalColor; // null = no face-level override was present
   }
   ```

2. In `ApplyToBody()` — accept a `List<SavedFaceColor>` parameter. Before setting each face color, read and save the original:
   ```csharp
   double[] original = face.MaterialPropertyValues as double[];
   savedList.Add(new SavedFaceColor {
       Face = face,
       OriginalColor = original != null ? (double[])original.Clone() : null
   });
   ```

3. Thread the saved list through `ApplyToPart()` → `ApplyToBody()` and `ApplyToAssembly()` → `ApplyToComponentTree()` → `ApplyToBody()`. Store in `_savedColors[path]` after applying.

4. Replace `ClearBodyColors()` with a new `RestoreSavedColors(string path)` method that:
   - Looks up saved colors for the model path
   - For each saved face: restores original value (null → set null, non-null → set original array)
   - Does **NOT** call `body.RemoveMaterialProperty()` (that was destructive)
   - Falls back to the current null-clearing approach if no saved colors exist (defensive)
   - Cleans up the saved colors entry from the dictionary

5. Update `RemoveFromPart()` and `RemoveFromAssembly()` to call `RestoreSavedColors()` instead.

### Step 2: Handle assembly components properly

For assemblies, component faces are accessed through the component's model doc. The save/restore keys on model path, and since all instances of a part share the same IPartDoc, one save/restore per unique part file is sufficient.

### Step 3: Create SideIndicatorQA test class

**New file:** `src/NM.SwAddin/Pipeline/SideIndicatorQA.cs`

**Test sequence (3-state validation):**
1. **Open** a sheet metal test part (B1_NativeBracket_14ga_CS.SLDPRT)
2. **STATE 1 — Baseline:** Read `face.MaterialPropertyValues` for every face on every body. Record as `baselineColors[]` (array of face-index → color-or-null).
3. **ACTION 1 — Apply:** Call `SideIndicatorService.Toggle()` (turns ON)
4. **STATE 2 — Applied:** Read all face colors again. Verify:
   - At least one face has green color (R≈0, G≈0.8, B≈0)
   - At least one face has red color (R≈0.8, G≈0, B≈0)
   - At least one face has gray color (R≈0.7, G≈0.7, B≈0.7)
   - No face matches baseline (all should be overridden)
5. **ACTION 2 — Remove:** Call `SideIndicatorService.Toggle()` (turns OFF)
6. **STATE 3 — Restored:** Read all face colors again. Verify:
   - Each face color matches its baseline value (null == null, or arrays match element-by-element within tolerance)
   - The service reports `IsActive == false`
7. **Report** pass/fail with detailed color comparison output

**Additional test cases:**
- **Test on B2_ImportedBracket** (imported sheet metal — uses fallback normal detection)
- **Idempotency test:** apply → remove → apply → remove (verify restore works across multiple cycles)
- **Color comparison utility:** `ColorsMatch(double[] a, double[] b, double tolerance = 0.001)` — handles both-null, one-null, and element-wise comparison

### Step 4: Integrate into BatchRunner

**File:** `src/NM.BatchRunner/Program.cs`

Add a new command-line flag: `--side-indicator-qa`

```csharp
case "--side-indicator-qa":
    var siQA = new SideIndicatorQA();
    int siResult = siQA.Run(swApp, inputDir);
    Environment.Exit(siResult); // 0=pass, 1=fail
    break;
```

### Step 5: Build and verify

1. Run `sync-csproj.ps1` (new file SideIndicatorQA.cs)
2. Run `build-and-test.ps1` to verify compilation
3. Document how to run: `NM.BatchRunner.exe --side-indicator-qa`

---

## File Changes Summary

| File | Action | Description |
|------|--------|-------------|
| `src/NM.SwAddin/SheetMetal/SideIndicatorService.cs` | Modify | Add save/restore, remove destructive body.RemoveMaterialProperty |
| `src/NM.SwAddin/Pipeline/SideIndicatorQA.cs` | New | 3-state color validation test |
| `src/NM.BatchRunner/Program.cs` | Modify | Add `--side-indicator-qa` command |
| `swcsharpaddin.csproj` | Modify (via sync-csproj.ps1) | Include new .cs file |

## Execution

**Execution: Sequential** — all changes are interdependent (fix first, then test harness that exercises the fix).

## Success Criteria

1. Build succeeds (`build-and-test.ps1` → `BUILD: SUCCESS`)
2. No new warnings introduced
3. SideIndicatorQA validates the 3-state cycle:
   - Baseline colors captured
   - Green/red/gray applied correctly
   - Original colors restored exactly after toggle off
4. Toggle works correctly for:
   - Parts with no explicit face colors (most common case)
   - Parts with body-level colors
   - Multiple toggle cycles (idempotent)
