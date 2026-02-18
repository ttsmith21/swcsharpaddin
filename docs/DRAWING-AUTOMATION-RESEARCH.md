# SolidWorks Drawing Automation — Comprehensive Research Report

## Executive Summary

After exhaustive research into SolidWorks API improvements from 2010-2026, we found **3 dramatic improvements** that could significantly enhance our drawing automation for STEP-imported sheet metal parts:

| Improvement | Impact | Works on STEP? | Effort |
|-------------|--------|----------------|--------|
| **DimXpert Auto Dimension Scheme** | Automatic hole/slot/pocket/boss dimensioning | **YES** — recognizes geometry, not feature tree | Medium |
| **Hole Table Automation** (`InsertHoleTable2`) | Complete hole callout tables on flat patterns | **YES** — reads geometry | Low |
| **IView::ImportAnnotations** (SW 2025+) | Import DimXpert annotations into drawing views | **YES** — imports DimXpert results | Low (SW 2025 req.) |

The **hybrid approach** — combining DimXpert (for feature dimensions) with our existing bend-line dimensioning (PR #39) — would take our drawings from "overall W×H + bend distances" to **fully-dimensioned manufacturing drawings** with hole callouts, slot dimensions, and feature recognition.

---

## Year-by-Year API Improvements (Drawing Automation Focus)

### SolidWorks 2010-2012 — Foundation APIs

| API | Purpose | STEP-Compatible? |
|-----|---------|-----------------|
| `CreateFlatPatternViewFromModelView3` | Dedicated flat pattern view creation | Yes |
| `InsertModelAnnotations3` | Import model dimensions into drawings | No (needs sketches) |
| `DropDrawingViewFromPalette2` | Drop views from palette | Yes |
| `GenerateViewPaletteViews` | Generate palette views from model | Yes |
| `DimXpert` introduced (2008, matured 2012) | 3D auto-dimensioning on geometry | **Yes** |
| `InsertModelInPredefinedView` | Insert model into template views | Yes |

**Key insight:** `CreateFlatPatternViewFromModelView3` has been available since 2010 and is more reliable than palette drop for sheet metal. Our code on main still uses palette drop — this is a gap.

### SolidWorks 2013-2015 — DimXpert Maturation

| API | Purpose | STEP-Compatible? |
|-----|---------|-----------------|
| `DimXpertAutoDimSchemeOption` | Configure auto-dimension parameters | **Yes** |
| `IDimXpertPart.CreateAutoDimensionScheme` | Run auto-dimensioning on part | **Yes** |
| `IView.InsertHoleTable2` | Insert hole tables in drawing views | **Yes** |
| `GetPolylines7` | Get visible view geometry edges | Yes |
| `GetBendLines` | Get bend lines from flat pattern view | Yes |

**Key insight:** `InsertHoleTable2` has been available since ~2015. Our VBA code had `AddHoleTable` but the C# port on main does NOT use it.

### SolidWorks 2016-2017 — DimXpert Improvements

| Feature | Details |
|---------|---------|
| DimXpert for chamfers and fillets | Expanded geometric feature recognition |
| DimXpert pattern recognition | Recognizes bolt circle patterns, slot patterns |
| Improved tolerance schemes | Plus/minus and geometric tolerance options |

### SolidWorks 2018-2019 — Developer Tooling

| API | Purpose | Drawing Impact |
|-----|---------|---------------|
| VSTA 3.0 (VS 2015 macros) | Modern .NET macro development | Better debugging |
| Bounding Box feature data API | Access bounding box dimensions | Useful for auto-layout |
| Table anchor API | Set table position anchors | Better hole table placement |
| `IView.IsFlatPatternView` | Detect flat pattern views programmatically | Validation |

### SolidWorks 2020-2022 — Performance & Drawing Speed

| Feature | Details |
|---------|---------|
| Detailing Mode | Open drawings instantly for annotation work |
| Graphics Acceleration for Drawings | Faster pan/zoom in drawing views |
| `StartDrawing`/`EndDrawing` | Disable inferencing for faster geometry creation |
| Large Assembly Drawings mode | Work with complex assemblies faster |
| `CloseAndReopen2` with ExitDetailingMode | API to switch detailing modes |
| SOLIDWORKS Inspection API (2022) | Automated inspection dimensioning |
| `IView.GetVisibleDrawingComponents` (2021) | Get unobscured components |

### SolidWorks 2023-2024 — Sheet Metal API Enhancements

| Feature | Details |
|---------|---------|
| Sheet metal API enhancements (2023) | Many new sheet metal calls |
| Backwards-compatible save (2024) | Save 2024 files to 2022/2023 |

### SolidWorks 2025 — MAJOR: IView::ImportAnnotations

| API | Purpose | Impact |
|-----|---------|--------|
| **`IView::ImportAnnotations`** | Import DimXpert annotations directly into drawing views via API | **Game-changer** for DimXpert-to-Drawing pipeline |
| `InsertBomTable6` (replaces `InsertBomTable5`) | Updated BOM table insertion | Minor |
| `ReloadWithReferences` performance | Selective reference reloading | Performance |

**Key insight:** Before SW 2025, importing DimXpert annotations into drawings required UI workarounds (check "Import Annotations" + "DimXpert Annotations" in view properties). SW 2025 finally exposes this as an API call.

### SolidWorks 2026 — AI-Powered Drawing Automation (Beta)

| Feature | Details | API Available? |
|---------|---------|---------------|
| **Auto-Generate Drawing (Beta)** | AI-powered automatic drawing creation with views, dimensions | **No API yet** — UI only |
| Automatic hole recognition | Recognizes holes in native AND imported geometry | Part of Auto-Generate |
| Auto view arrangement | Arranges views to avoid overlap, matches sheet format | Part of Auto-Generate |

**Key insight:** SW 2026's Auto-Generate Drawing is exactly what we want — but it's Beta, UI-only, and has no API. When they expose it via API (likely SW 2027-2028), it could replace our entire drawing pipeline. For now, we build our own version using the existing APIs.

---

## The 3 Dramatic Improvements (Detail)

### 1. DimXpert Auto Dimension Scheme — Feature Recognition for STEP Files

**What it does:** DimXpert analyzes the solid body geometry (NOT the feature tree) and automatically recognizes manufacturing features: holes, slots, pockets, bosses, fillets, chamfers, shoulders, notches. It then applies dimensions and tolerances to these features.

**Why this is dramatic:** Our current approach (PR #39) only dimensions:
- Overall width x height
- Bend-line-to-bend-line distances
- Bend-line-to-edge distances

With DimXpert, we would ALSO get:
- Hole positions and diameters
- Slot width/length/position
- Pocket dimensions
- Boss height/diameter
- Fillet/chamfer callouts
- Pattern spacing (bolt circles, linear patterns)

**API Pattern (C#):**
```csharp
using SolidWorks.Interop.swdimxpert;

// 1. Get DimXpertManager from model extension
DimXpertManager swSchema = swModel.Extension.DimXpertManager["Default", true];
DimXpertPart swDXPart = swSchema.DimXpertPart;

// 2. Configure auto-dimension scheme
DimXpertAutoDimSchemeOption schemeOption = swDXPart.GetAutoDimSchemeOption();

// 3. Set datum features (3 datums: primary, secondary, tertiary)
// Select faces for datums using SelectByID2 with marks
schemeOption.PrimaryDatumFace = ...; // largest flat face
schemeOption.SecondaryDatumFace = ...; // perpendicular edge face
schemeOption.TertiaryDatumFace = ...; // third orthogonal face

// 4. Run auto-dimension
swDXPart.CreateAutoDimensionScheme(schemeOption);

// 5. Get results — features and annotations
object[] features = swDXPart.GetFeatures();
foreach (DimXpertFeature feat in features)
{
    swDimXpertFeatureType_e type = feat.GetFeatureType();
    // Types: Hole, Slot, Pocket, Boss, Fillet, Chamfer, etc.

    object[] annotations = feat.GetAppliedAnnotations();
    foreach (DimXpertAnnotation ann in annotations)
    {
        // Position, size, tolerance annotations
    }
}
```

**Critical detail for STEP imports:** DimXpert works because it recognizes **geometric/manufacturing features** from the B-Rep solid body, NOT from the FeatureManager Design Tree. A STEP import has a single "Imported" feature but DimXpert still finds all the holes, slots, pockets, etc. from the geometry.

**Integration with drawings:** After running DimXpert on the 3D part:
- **SW 2025+:** Use `IView.ImportAnnotations` to pull DimXpert dims into drawing views
- **SW 2022-2024:** Use Windows API workaround (CodeStack pattern) to check "Import Annotations" + "DimXpert Annotations" checkboxes on the view PropertyManager
- **Alternative:** Read DimXpert results programmatically and place equivalent `AddDimension2` calls manually in the drawing view

### 2. Hole Table Automation (InsertHoleTable2)

**What it does:** Automatically generates a table listing every hole in the flat pattern view with:
- Hole tag (A1, A2, B1, etc.)
- X/Y position from datum
- Hole diameter
- Hole depth
- Hole count

**Why this is dramatic:** Our VBA code had `AddHoleTable` but the C# port on main does NOT implement it. For shop floor use, a hole table is essential — the press brake operator needs to know where to punch/drill.

**API Pattern (C#):**
```csharp
// 1. Pre-select required entities on the flat pattern view
// Mark 1 = datum vertex (origin point)
// Mark 2 = holes face (flat pattern face)
// Mark 4 = X-axis edge
// Mark 8 = Y-axis edge

swModel.Extension.SelectByID2("", "VERTEX", x, y, z, false, 1, null, 0);
swModel.Extension.SelectByID2("", "FACE", x, y, z, true, 2, null, 0);
swModel.Extension.SelectByID2("", "EDGE", x, y, z, true, 4, null, 0);
swModel.Extension.SelectByID2("", "EDGE", x, y, z, true, 8, null, 0);

// 2. Insert hole table
IView flatPatternView = ...; // our flat pattern view
HoleTableAnnotation holeTable = flatPatternView.InsertHoleTable2(
    false,          // UseAnchorPoint
    0.25,           // X position (meters)
    0.01,           // Y position (meters)
    (int)swBomBalloonStyle_e.swBS_Circular, // Balloon style
    "",             // Hole table template (empty = default)
    ""              // Origin indicator
);
```

### 3. IView::ImportAnnotations (SW 2025+)

**What it does:** Programmatically imports DimXpert annotations from the 3D model into a drawing view — the API equivalent of checking the "Import Annotations" + "DimXpert Annotations" checkboxes manually.

**Why this is dramatic:** Before SW 2025, there was NO API to import DimXpert annotations into drawings. The only options were:
1. Use `InsertModelAnnotations3` (doesn't handle DimXpert)
2. Use Windows API to automate the UI PropertyManager (fragile hack)
3. Read DimXpert data and manually re-create dimensions (complex)

With SW 2025, it's a single API call.

---

## Recommended Hybrid Approach

Combine our existing bend-line dimensioning (PR #39) with DimXpert for a dramatically better result:

### Pipeline (per part):

```
1. Open STEP file in SolidWorks
2. Run InsertBends2 -> flatten -> get flat pattern
3. [NEW] Run DimXpert Auto Dimension Scheme on 3D part
   -> Recognizes holes, slots, pockets, bosses automatically
4. Create drawing from template
5. Drop flat pattern view
6. [NEW] Import DimXpert annotations into view (SW 2025)
   OR manually place DimXpert-derived dims (SW 2022)
7. [EXISTING] Run DimensionFlatPattern() for overall W x H + bend dims
8. [NEW] Insert hole table via InsertHoleTable2
9. [EXISTING] AlignDimensions (auto-arrange)
10. [EXISTING] Validate (dangling dims, undimensioned bends)
11. Save drawing + export DXF
```

### Result: Before vs After

| Dimension Type | Current (PR #39) | With DimXpert Hybrid |
|---------------|-----------------|---------------------|
| Overall Width x Height | Yes | Yes |
| Bend-to-bend distances | Yes | Yes |
| Bend-to-edge distances | Yes | Yes |
| **Hole positions** | No | **Yes** |
| **Hole diameters** | No | **Yes** |
| **Slot dimensions** | No | **Yes** |
| **Pocket dimensions** | No | **Yes** |
| **Hole table** | No | **Yes** |
| **Pattern spacing** | No | **Yes** |
| **Feature callouts** | No | **Yes** |

---

## All Gaps

### Gap 1: DimXpert Auto Dimensioning (High Impact)

Run DimXpert on 3D part before drawing creation to automatically dimension holes, slots, pockets, bosses. Works on STEP imports.

**Implementation:** New `DimXpertDimensioner.cs` in `src/NM.SwAddin/Drawing/`
- Access `Extension.DimXpertManager["Default", true]`
- Configure `DimXpertAutoDimSchemeOption` with datum faces
- Call `CreateAutoDimensionScheme()`
- Import results into drawing view

**Complexity:** Medium — need to auto-select datum faces (largest flat face + 2 perpendicular edges)

### Gap 2: Hole Table Automation (High Impact)

Insert hole tables on flat pattern views automatically.

**Implementation:** Add `InsertHoleTable()` method to `DrawingDimensioner.cs`
- Pre-select datum vertex, face, X/Y edges with marks
- Call `IView.InsertHoleTable2()`
- Position table below or beside the flat pattern view

**Complexity:** Low-Medium

### Gap 3: IView::ImportAnnotations (Medium Impact, SW 2025+ only)

Use the new SW 2025 API to import DimXpert annotations into drawing views.

**Implementation:** Single API call after view creation
**Complexity:** Low — but requires SW 2025.

### Gap 4: CreateFlatPatternViewFromModelView3 (Medium Impact)

Replace `DropDrawingViewFromPalette2("Flat Pattern", ...)` with the dedicated flat pattern view API.

**Complexity:** Low

### Gap 5: Dual DXF Export (Moderate Impact)

Export `_CUT.dxf` (geometry only) and `_FULL.dxf` (geometry + bend lines + sketches).

**Complexity:** Low

### Gap 6: DXF Validation (Low Impact)

Validate exported DXFs are non-empty (file exists, size > 100 bytes).

**Complexity:** Very Low

---

## Priority Ranking

| Priority | Gap | Impact | Effort | Recommendation |
|----------|-----|--------|--------|---------------|
| **P0** | DimXpert Auto Dimensioning | Very High | Medium | Implement first |
| **P0** | Hole Table Automation | High | Low-Medium | Implement with DimXpert |
| **P1** | ImportAnnotations (SW 2025) | Medium | Low | Add when on SW 2025 |
| **P1** | CreateFlatPatternViewFromModelView3 | Medium | Low | Quick win |
| **P2** | DXF Validation | Low | Very Low | Safety net |
| **P2** | Dual DXF Export | Moderate | Low | Shop floor value |

---

## Key References

### API Documentation
- [DimXpert Auto Dimension Scheme Example (C#)](https://help.solidworks.com/2024/English/api/swdimxpertapi/Auto_Dimension_Scheme_Example_CSharp.htm)
- [DimXpert Features and Annotations Example (C#)](https://help.solidworks.com/2024/english/api/swdimxpertapi/Get_DimXpert_Features_and_Annotations_in_a_Model_Example_CSharp.htm)
- [InsertHoleTable2 Method (IView)](https://help.solidworks.com/2022/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IView~InsertHoleTable2.html)
- [CreateFlatPatternViewFromModelView3](https://help.solidworks.com/2022/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IDrawingDoc~CreateFlatPatternViewFromModelView3.html)
- [InsertModelAnnotations3 Method](https://help.solidworks.com/2022/english/api/sldworksapi/solidworks.interop.sldworks~solidworks.interop.sldworks.idrawingdoc~insertmodelannotations3.html)
- [DimXpert Dimensions and Drawings](https://help.solidworks.com/2024/english/SolidWorks/sldworks/c_dimxpert_dimensions_and_drawings.htm)

### Year-by-Year API Changes (CADSharp)
- [What's New in the 2018 SOLIDWORKS API](https://www.cadsharp.com/blog/whats-new-in-the-solidworks-2018-api/)
- [What's New in the 2019 SOLIDWORKS API](https://www.cadsharp.com/blog/whats-new-2019-api/)
- [What's New in the 2024 SOLIDWORKS API](https://www.cadsharp.com/blog/whats-new-2024-api/)
- [What's New in the 2025 SOLIDWORKS API](https://www.cadsharp.com/blog/whats-new-2025-solidworks-api/)
- [What's New in the 2026 SOLIDWORKS API](https://www.cadsharp.com/blog/whats-new-in-the-2026-solidworks-api/)

### DimXpert Guides
- [Introduction to DimXpert — GoEngineer](https://www.goengineer.com/blog/introduction-dimxpert-solidworks)
- [DimXpert Features — SW 2022 Help](https://help.solidworks.com/2022/english/solidworks/sldworks/c_dimxpert_features.htm)
- [MBD, DimXpert, and MBD Dimensions — Hawk Ridge](https://hawkridgesys.com/blog/mbd-dimxpert-and-mbd-dimensions-whats-the-difference)
- [Importing DimXpert Into Drawings — CATI](https://www.cati.com/blog/solidworks-importing-dimxpert-into-drawings/)

### Drawing Automation Resources
- [CodeStack — Drawing Automation](https://www.codestack.net/solidworks-api/document/drawing/)
- [CodeStack — Insert Holes Table Macro](https://www.codestack.net/solidworks-api/document/tables/insert-holes-table/)
- [Eng-Tips — Flat Pattern Drawing Automation](https://www.eng-tips.com/threads/solidworks-automation-creating-a-flat-pattern-drawing-from-a-sheet-metal-part-with-automatic-dims.499245/)

### SolidWorks What's New (Official)
- [What's New in SOLIDWORKS 2026 — Design](https://blogs.solidworks.com/solidworksblog/2025/10/whats-new-in-solidworks-2026-design.html)
- [What's New in SOLIDWORKS 2026 — Official](https://www.solidworks.com/media/whats-new-solidworks-2026)
- [SW 2025 API Help — What's New](https://help.solidworks.com/2025/english/WhatsNew/c_wn2025_manage_api.htm)
