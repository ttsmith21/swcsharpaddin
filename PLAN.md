# Rename Wizard - Implementation Plan

## Overview

Add a "Rename Wizard" button to the add-in that enables AI-powered batch renaming of components in an assembly. The wizard extracts a BOM from a customer PDF drawing, matches BOM rows to STEP-imported component names using AI, and presents an editable grid for the user to approve/modify renames. Actual file rename uses the SolidWorks `RenameDocument` API (in-memory rename + automatic reference updates).

## Architecture

```
New button: "Rename Wizard"
         │
         ▼
  RenameWizardRunner (orchestrator)
    ├── PdfDrawingAnalyzer          ← REUSE (extract BOM from PDF)
    ├── VisionPrompts.GetBomTablePrompt()  ← NEW prompt
    ├── ComponentCollector          ← REUSE (walk assembly tree)
    ├── BomComponentMatcher         ← NEW (AI matches BOM→components)
    ├── RenameWizardForm            ← NEW (DataGridView UI)
    └── AssemblyRenameService       ← NEW (RenameDocument API)
```

## Execution: Sequential

This is a single feature with dependent steps. No worktrees needed.

---

## Step 1: Data Models (`src/NM.Core/Rename/`)

Create new folder `src/NM.Core/Rename/` with:

### `BomRow.cs`
```csharp
public sealed class BomRow
{
    public int ItemNumber { get; set; }
    public string PartNumber { get; set; }
    public string Description { get; set; }
    public string Material { get; set; }
    public int Quantity { get; set; }
}
```

### `RenameEntry.cs`
```csharp
public sealed class RenameEntry
{
    public int Index { get; set; }
    public string CurrentFileName { get; set; }   // e.g. "Body-Move-Copy1.SLDPRT"
    public string CurrentFilePath { get; set; }    // full path
    public string PredictedName { get; set; }      // AI suggestion
    public string FinalName { get; set; }          // user-editable (starts = predicted)
    public double Confidence { get; set; }         // 0.0 to 1.0
    public string MatchReason { get; set; }        // "BOM item 3: Bracket-Left"
    public BomRow MatchedBomRow { get; set; }      // null if no match
    public IComponent2 Component { get; set; }     // for click-to-highlight
    public string Configuration { get; set; }
    public bool IsApproved { get; set; }           // user checks/unchecks
}
```

### `RenameWizardResult.cs`
```csharp
public sealed class RenameWizardResult
{
    public int TotalComponents { get; set; }
    public int Renamed { get; set; }
    public int Skipped { get; set; }
    public int Failed { get; set; }
    public List<string> Errors { get; set; } = new List<string>();
    public string Summary { get; set; }
}
```

---

## Step 2: BOM Table Vision Prompt (`src/NM.Core/AI/VisionPrompts.cs`)

Add new method `GetBomTablePrompt()` that asks Claude to extract a structured BOM table from an assembly drawing page. Returns JSON array of items with part_number, description, material, quantity. Follow existing prompt pattern (JSON schema + RULES section).

---

## Step 3: BOM Component Matcher (`src/NM.Core/Rename/BomComponentMatcher.cs`)

**Purpose:** Match BOM rows from the PDF to assembly components from the STEP import.

**Matching strategy (layered):**

1. **Exact name match** — BOM part number appears in STEP filename (highest confidence)
2. **Fuzzy name match** — Levenshtein/substring similarity between BOM description and component name
3. **Material correlation** — If materials already processed, match by material type
4. **Quantity match** — BOM qty matches component instance count
5. **AI fallback** — Send both lists to Claude and ask it to match them (expensive, last resort)

**Input:** `List<BomRow>` + `List<ComponentInfo>` (path, name, material, qty)
**Output:** `List<RenameEntry>` with predictions and confidence scores

For AI fallback, add a `GetBomMatchingPrompt()` to VisionPrompts that sends both lists as JSON and asks Claude to return the mapping.

---

## Step 4: Assembly Rename Service (`src/NM.SwAddin/Rename/AssemblyRenameService.cs`)

**Purpose:** Execute approved renames using SolidWorks `RenameDocument` API.

**Algorithm:**
```
For each approved RenameEntry:
  1. Select component in assembly: comp.Select4(false, null, false)
  2. Call ext.RenameDocument(newNameWithoutExtension)
  3. Check return: swRenameDocumentError_e.swRenameDocumentError_None
  4. Track success/failure

After all renames:
  5. Get IRenamedDocumentReferences via ext.GetRenamedDocumentReferences()
  6. Set UpdateWhereUsedReferences = true
  7. Call Search() to find/update unopened docs that reference old names
  8. Save assembly: model.Save3(...)
  9. Save each renamed component
```

**Error handling:**
- If a single rename fails, log it, skip it, continue with others
- Report all failures in the result summary
- If ALL renames fail, show error and don't attempt save

**Validation (pre-rename):**
- Check no two components map to the same new name (collision detection)
- Check new names are valid filenames (no illegal chars)
- Check destination files don't already exist on disk
- Reuse `FileRenameValidator.IsValidFileName()` pattern

---

## Step 5: Rename Wizard Form (`src/NM.SwAddin/UI/RenameWizardForm.cs`)

**WinForms dialog, follows PropertyReviewWizard pattern.**

**Size:** 960 x 700, resizable

**Layout:**
```
┌─────────────────────────────────────────────────┐
│ Assembly: TopAssembly.SLDASM                     │
│ Drawing:  CustomerDrawing.pdf                    │
│ Components: 23 unique  |  BOM items: 25          │
├─────────────────────────────────────────────────┤
│ Find: [________] Replace: [________] [Apply]     │
├─┬──────────────┬──────────────┬─────────────┬───┤
│✓│ Current Name │ AI Predicted │ New Name    │Cnf│
├─┼──────────────┼──────────────┼─────────────┼───┤
│☑│ Body-Move1   │ Bracket-Left │ Bracket-Left│95%│
│☑│ Import-47    │ Base-Plate   │ Base-Plate  │88%│
│☐│ Part3^Assy   │ (no match)   │ Part3^Assy  │ 0%│
│  │ ...          │              │             │   │
├─┴──────────────┴──────────────┴─────────────┴───┤
│ [Select All] [Deselect All]    [Rename] [Cancel] │
└─────────────────────────────────────────────────┘
```

**Columns:**
1. **Checkbox** — `DataGridViewCheckBoxColumn`, auto-checked if confidence >= 0.80
2. **Current Name** — read-only, filename without extension
3. **AI Prediction** — read-only, gray text if no match
4. **New Name** — **editable** `DataGridViewTextBoxColumn` (starts with AI prediction or current name)
5. **Confidence** — read-only, color-coded (green >=80%, yellow 50-80%, red <50%)

**Features:**
- **Find/Replace bar:** Two textboxes + "Apply" button. On click, iterates all rows and does string.Replace on the "New Name" column. Supports multiple applications (not regex, just simple string replace).
- **Click-to-highlight:** On row selection changed, call `component.Select4(true, null, false)` then `model.ViewZoomToSelection()` to highlight and zoom to the selected component in the SolidWorks viewport.
- **Color coding:** Green row = high confidence match. Yellow = medium. White = no match (user must fill in or skip).
- **Validation on Apply:** Check for duplicate new names, empty names, invalid characters. Show warning if issues found.

**Public properties (read after ShowDialog):**
- `ApprovedRenames: List<RenameEntry>` — entries where checkbox is checked
- `DialogResult` — OK or Cancel

---

## Step 6: Rename Wizard Runner (`src/NM.SwAddin/Pipeline/RenameWizardRunner.cs`)

**Orchestrator, follows DrawingAnalysisRunner pattern.**

**Constructor:** `RenameWizardRunner(ISldWorks swApp)`

**Main method: `RunOnActiveAssembly()`**

**Pipeline:**
```
1. Validate active doc is assembly
     → If not, show "Open an assembly first" and return
2. Collect components via ComponentCollector
     → Get unique parts with paths, names, configs
     → Also get quantity per component
3. Prompt user to select PDF (reuse PdfDrawingAnalyzer.FindCompanionPdf + OpenFileDialog)
4. Extract BOM from PDF
     → PdfDrawingAnalyzer with new GetBomTablePrompt()
     → Parse JSON response into List<BomRow>
5. Match BOM rows to components
     → BomComponentMatcher.Match(bomRows, components)
     → Returns List<RenameEntry> with predictions
6. Show RenameWizardForm
     → Pass List<RenameEntry>
     → User reviews, edits names, uses find/replace
     → Returns approved renames on OK, null on Cancel
7. Validate approved renames
     → FileRenameValidator for each entry
     → Check for name collisions
     → Show warnings if any, let user proceed or cancel
8. Execute renames
     → AssemblyRenameService.RenameComponents(model, approvedRenames)
     → Returns RenameWizardResult
9. Show summary
     → "Renamed 18 of 23 components. 2 skipped, 3 failed."
     → List any errors
```

---

## Step 7: Button Registration (`swaddin.cs`)

1. Add constant: `public const int mainItemID11 = 10;  // Rename Wizard`
2. Add to `knownIDs` arrays (both DEBUG and RELEASE)
3. Add `AddCommandItem2("Rename Wizard", ...)` with tooltip "AI-powered batch rename of assembly components using PDF drawing BOM"
4. Use next available icon index (7 or create new icon)
5. Add callback method: `public void RenameWizard()` following AnalyzeDrawing() pattern
6. Add to command tab arrays
7. Enable callback: `public int RenameWizardEnable()` — return enabled only if active doc is assembly

---

## Step 8: Wire Up & Sync

1. Run `sync-csproj.ps1` to add all new .cs files to the csproj
2. Run `build-and-test.ps1` to verify compilation
3. Fix any CS0246/namespace issues

---

## New Files Summary

| File | Location | Type |
|------|----------|------|
| `BomRow.cs` | `src/NM.Core/Rename/` | Model |
| `RenameEntry.cs` | `src/NM.Core/Rename/` | Model |
| `RenameWizardResult.cs` | `src/NM.Core/Rename/` | Model |
| `BomComponentMatcher.cs` | `src/NM.Core/Rename/` | Logic |
| `AssemblyRenameService.cs` | `src/NM.SwAddin/Rename/` | SW API |
| `RenameWizardForm.cs` | `src/NM.SwAddin/UI/` | WinForms |
| `RenameWizardRunner.cs` | `src/NM.SwAddin/Pipeline/` | Orchestrator |

**Modified files:**
- `swaddin.cs` — button registration + callback
- `src/NM.Core/AI/VisionPrompts.cs` — new BOM extraction prompt
- `swcsharpaddin.csproj` — new file references (via sync-csproj.ps1)

---

## Risk Mitigation

| Risk | Mitigation |
|------|------------|
| `RenameDocument` fails on some components | Log error, skip that component, continue with others. Show failures in summary. |
| BOM extraction misses items | User can manually type new names in the editable column. AI is a starting point, not mandatory. |
| AI matching is wrong | Color-coded confidence + click-to-highlight lets user verify before approving |
| Name collision (two parts → same new name) | Pre-validate before executing. Block rename if collision detected. |
| Companion drawings (.SLDDRW) not updated | `IRenamedDocumentReferences.Search()` handles this if drawings are in same folder |
| User renames to invalid filename | Validate with `FileRenameValidator.IsValidFileName()` before executing |

---

## Testing Strategy

- **Smoke tests:** BomComponentMatcher matching logic (pure C#, no SW needed)
- **Manual QA:** Open test assembly with STEP-imported parts, run wizard, verify renames stick after save/close/reopen
- **Build verification:** `build-and-test.ps1` must pass with no new errors
