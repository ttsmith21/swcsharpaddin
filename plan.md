# Plan: Problem Parts Skip, Suppression Audit, and Toggle-Red Button

## Item 1: Fix — Already-classified parts should be skipped on re-run

**Problem:** When a surface body is marked PUR on the first run, the second run still flags it as a problem because `PartPreflight` rejects it on geometry (no solid body) before `MainRunner` ever checks `rbPartType`.

**Approach:** Add an early-out check in `BatchValidator.ValidateSingleModel()` that reads `rbPartType` from the file's custom properties. If `rbPartType == "1"`, skip validation entirely and route to GoodModels. MainRunner already handles the processing early-out.

**Files to modify:**
- `src/NM.SwAddin/Validation/BatchValidator.cs` — Add rbPartType check in ValidateSingleModel, before calling `_validator.Validate()`

**Details:**
1. After opening the document (line ~144 in BatchValidator.cs), read `rbPartType` custom property
2. If `rbPartType == "1"`, mark the model as validated (good), add to GoodModels, and return early
3. Log a debug message: `[BATCHVAL] rbPartType=1 detected - skipping validation (already classified)`
4. This lets MainRunner handle the PUR/MACH/CUST early-out during processing

**Risk:** Low. This is a read-only check on a custom property. Parts without the property are unaffected.

**Execution: Sequential**

---

## Item 2: Audit — Suppressed parts behavior (no code changes needed)

**Finding:** Suppressed parts are already correctly ignored at three layers:
1. `GetComponents(false)` excludes suppressed from the API array
2. `ComponentValidator` explicitly rejects suppressed components
3. `AssemblyComponentQuantifier` skips suppressed in recursive traversal

**No code changes needed.** The current implementation matches the expected behavior.

---

## Item 3: Feature — Toggle problem parts red on toolbar

**Approach:** Add a new toolbar button "Toggle Problem Colors" that:
- Reads problem parts from `ProblemPartManager.Instance`
- For each component in the active assembly, if its file path matches a problem part, sets its appearance to red
- On second click (toggle off), restores original appearance
- Uses enable callback to only be active when an assembly is open

**Files to create/modify:**
1. `src/NM.SwAddin/UI/ProblemPartColorizer.cs` (NEW) — Logic for applying/removing red color to components
2. `SwAddin.cs` — Register new command, add callback method
3. Icon resources — Add icon to toolbar strip PNGs (or reuse existing index)

**Details for ProblemPartColorizer:**
```
public sealed class ProblemPartColorizer
{
    private bool _isActive;
    private Dictionary<string, double[]> _savedAppearances; // path → original values

    public void Toggle(ISldWorks swApp)
    {
        if (_isActive) RemoveColors(swApp);
        else ApplyColors(swApp);
        _isActive = !_isActive;
    }

    private void ApplyColors(ISldWorks swApp)
    {
        // Get active assembly
        // Get all components via GetComponents(false)
        // For each component, check if file path is in ProblemPartManager
        // If yes: save current MaterialPropertyValues, then set to red
        // Red RGB = [1.0, 0.0, 0.0] with ambient/diffuse/specular/emissive coefficients
    }

    private void RemoveColors(ISldWorks swApp)
    {
        // Restore saved appearances from _savedAppearances dictionary
    }
}
```

**SwAddin.cs changes:**
1. Add `mainItemID9 = 8` constant
2. Add to `knownIDs` array
3. Register via `AddCommandItem2("Toggle Problem Colors", ...)`
4. Add `ToggleProblemColors()` callback method
5. Add `ToggleProblemColorsEnable()` — return 1 only if assembly is active and problem parts exist

**SolidWorks API for component color:**
```csharp
// Get current appearance
double[] props = (double[])component.GetMaterialPropertyValues2(
    (int)swInConfigurationOpts_e.swThisConfiguration, null);

// Set red appearance
// Array: [R, G, B, Ambient, Diffuse, Specular, Shininess, Transparency, Emission]
double[] red = new double[] { 1.0, 0.0, 0.0, 0.5, 1.0, 0.5, 0.5, 0.0, 0.0 };
component.SetMaterialPropertyValues2(red,
    (int)swInConfigurationOpts_e.swThisConfiguration, null);
```

**Risk:** Medium. Color manipulation is cosmetic and reversible. The saved-appearances dictionary handles restoration. Edge case: if user closes/reopens assembly, colors persist in session but _savedAppearances is lost — acceptable since the colors are visual-only and don't affect the model file.

**Execution: Sequential** (Item 1 first, then Item 3 — they share ProblemPartManager context)
