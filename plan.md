# Plan: Automated Drawing Creation with Dimensioning & Validation

## Overview

Port the VBA automated drawing creation system (`SP.bas::CreateDrawing`, `DimensionDrawing.bas`, `AlignDims.bas`) to C#. The existing `DrawingGenerator.cs` has a partial scaffold but is **missing** the critical auto-dimensioning, bend line dimensioning, tube dimensioning, view alignment, formed view creation, etch mark visibility, and drawing validation features.

Additionally, add two new validation checks:
- **Dangling dimension detection** (via color check strategy)
- **Bend line dimension completeness** (warn if any bend lines lack dimensions)

---

## VBA Code Analysis (What We Must Replicate)

### SP.bas::CreateDrawing() — Main Orchestrator (lines 1639-1979)

**Step-by-step flow:**
1. Save the model
2. Determine paths (drawing = `<partname>.slddrw`, DXF = `<partname>.dxf`)
3. Check if drawing already exists → skip if so
4. Read `OP20` property to determine part type
5. Read `PressBrake` and `F325` properties
6. **If sheet metal** (OP20 starts with N115/N120/N125):
   - Get/create FlatPatternFolder
   - Unsuppress flat pattern + sub-features
   - Create `DefaultSM-FLAT-PATTERN` configuration
   - Switch back to Default, suppress flat pattern in Default config
7. Create new drawing from template (`A-SIZE.drwdot`)
8. Generate view palette: `GenerateViewPaletteViews()`
9. Drop "Flat Pattern" view (fall back to "*Right" if no flat pattern)
10. Force rebuild
11. **Rotate view if taller than wide** (90 degrees)
12. **Check grain direction** — if either dimension < 6", set Grain="Y"
13. **Position view** to bottom-left with margins
14. **Clamp view within sheet bounds** (check outline vs 0.268 and 0.21)
15. Save DXF and/or drawing
16. **If TUBE LASER** (OP20 = "F110 - TUBE LASER"):
    - Reposition flat pattern view
    - `DimensionTube()` on flat pattern view
    - Create unfolded (formed) view via `CreateUnfoldedViewAt3()`
    - Set formed view to Default config
    - `DimensionTube()` on formed view
    - `AlignDims.Align()` on drawing
    - Auto-scale to fit both views on sheet
17. **Else (standard sheet metal)**:
    - `DimensionFlat()` on flat pattern view
    - `AlignDims.Align()` on drawing
18. **Hide reference planes** in the part
19. **Make etch marks visible** — find ProfileFeature sketches, unblank in drawing
20. Final save

### DimensionDrawing.bas::DimensionFlat() — Sheet Metal Auto-Dimensioning

**Data types:**
- `ProcessedElement` — line with X1,Y1,X2,Y2 and angle
- `BendElement` — bend line sketch segment with Position, P2, Angle
- `EdgeElement` — edge object with X,Y coordinates and type (Line/Point)

**Algorithm:**
1. Get view transform (`GetXform`)
2. `FindBendLines()` — uses `IView.GetBendLines()` to find horizontal and vertical bend lines, sort by position
3. Find bounding edges: `FindLeftPosLine`, `FindRightPosLine`, `FindTopPosLine`, `FindBottomPosLine` — traverse `GetPolylines7(1, Null)` to find extreme edges/vertices
4. If no bounding edges found → fall back to `DimensionTube()` (scale to fit + basic dims)
5. **Add overall horizontal dimension** (select left edge + right edge → `AddHorizontalDimension2`)
6. **Add overall vertical dimension** (select top edge + bottom edge → `AddVerticalDimension2`)
7. **Dimension between vertical bend lines** — iterate pairs, select EXTSKETCHSEGMENT by constructed name, `AddHorizontalDimension2`, `EditDimensionProperties2` with " BL" suffix
8. **Create top projected view** — `CreateUnfoldedViewAt3()`, set to Default config, hidden display, `DimensionOther()` for overall width/height
9. **Dimension between horizontal bend lines** — same pattern with `AddVerticalDimension2` and " BL" suffix
10. **Create right projected view** — same pattern
11. **Auto-scale** views to fit on sheet (iterative 5% scale-up loop with timeout of 25)
12. `AlignDims.Align()` — auto-arrange all dimensions

### DimensionDrawing.bas::DimensionTube() — Tube Auto-Dimensioning

1. Get view outline
2. Calculate edge positions from outline
3. Try to select left edge via `SelectByRay`, check if it's a circle (arc)
4. If circular: `AddDiameterDimension2()`
5. If rectangular: `RectProfile()` — select opposing edges, add horizontal + vertical dimensions
6. If neither works: try selecting bottom edge for horizontal dim

### AlignDims.bas::Align() — Dimension Auto-Arrangement

1. Iterate all views on the current sheet
2. For each view, get annotations
3. Select all dimension annotations
4. Call `Extension.AlignDimensions(swAlignDimensionType_AutoArrange, 0.001)`

### Key SolidWorks API Calls Used

| API Call | Purpose |
|----------|---------|
| `NewDocument(template, paperSize, width, height)` | Create drawing |
| `GenerateViewPaletteViews(modelPath)` | Generate view palette |
| `DropDrawingViewFromPalette2(viewName, x, y, z)` | Drop view onto sheet |
| `CreateUnfoldedViewAt3(x, y, z, false)` | Create projected/unfolded view |
| `DrawingViewRotate(radians)` | Rotate a view |
| `IView.GetOutline()` | Get view bounding box [x1,y1,x2,y2] |
| `IView.Position` | Get/set view position [x,y] |
| `IView.GetXform()` | View transform (scale at index 2) |
| `IView.GetBendLines()` | Get bend line sketch segments |
| `IView.GetPolylines7(1, Null)` | Get visible edges in view |
| `IView.SetDisplayMode3()` | Set view display mode (hidden, wireframe, etc.) |
| `IView.AlignWithView()` | Align projected view with parent |
| `IView.ReferencedConfiguration` | Set which config the view shows |
| `ISheet.GetProperties2()` | Sheet properties (scale at indices 2,3) |
| `ISheet.SetProperties2(...)` | Update sheet scale |
| `AddHorizontalDimension2(x, y, z)` | Add horizontal dimension |
| `AddVerticalDimension2(x, y, z)` | Add vertical dimension |
| `AddDiameterDimension2(x, y, z)` | Add diameter dimension |
| `EditDimensionProperties2(...)` | Set dimension properties (tolerance, suffix " BL") |
| `Extension.AlignDimensions(type, gap)` | Auto-arrange dimensions |
| `Extension.SelectByID2(name, type, ...)` | Select entity by name |
| `Extension.SelectByRay(...)` | Select by ray cast |
| `SketchSegment.GetName()` | Get bend line name for selection |
| `SketchSegment.GetSketch().Name` | Get parent sketch name |
| `IView.RootDrawingComponent.Name` | Get drawing component name |

---

## Implementation Plan

### Phase 1: Cache VBA Reference Files + Branch Setup

**Files to add:**
- `docs/vba-reference/DimensionDrawing.bas` (already cached)
- `docs/vba-reference/AlignDims.bas` (already cached)

### Phase 2: Core Data Types (NM.Core)

**New file:** `src/NM.Core/Drawing/DrawingTypes.cs`

Port the VBA types to C# structs/classes:

```csharp
namespace NM.Core.Drawing
{
    public class ProcessedElement
    {
        public object Obj { get; set; }
        public double X1, X2, Y1, Y2;
        public double Angle;
    }

    public class BendElement
    {
        public object SketchSegment { get; set; } // ISketchSegment
        public double Position;
        public double P2;
        public double Angle;
    }

    public class EdgeElement
    {
        public object Obj { get; set; }
        public double X, Y;
        public double Angle;
        public string Type; // "Line" or "Point"
    }

    public class DrawingValidationResult
    {
        public bool Success { get; set; }
        public List<string> DanglingDimensions { get; set; } = new List<string>();
        public List<string> UndimensionedBendLines { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
    }
}
```

### Phase 3: Geometry Analysis (NM.SwAddin)

**New file:** `src/NM.SwAddin/Drawing/ViewGeometryAnalyzer.cs`

Port the edge/line/bend finding functions:

| VBA Function | C# Method | Purpose |
|---|---|---|
| `ProcessLine()` | `ProcessLine()` | Transform line endpoints to view space, compute angle |
| `TransformSketchPointToModelSpace()` | `TransformSketchPoint()` | Transform sketch point through sketch→model transform |
| `FindBendLines()` | `FindBendLines()` | Get horizontal/vertical bend lines from `IView.GetBendLines()` |
| `FindLeftPosLine()` | `FindBoundaryEdge(Direction.Left)` | Find leftmost edge via `GetPolylines7` |
| `FindRightPosLine()` | `FindBoundaryEdge(Direction.Right)` | Find rightmost edge |
| `FindTopPosLine()` | `FindBoundaryEdge(Direction.Top)` | Find topmost edge |
| `FindBottomPosLine()` | `FindBoundaryEdge(Direction.Bottom)` | Find bottommost edge |
| `SortInfo()` | `SortBendElements()` | Sort bends by position, then P2 |

### Phase 4: Dimensioning Service (NM.SwAddin)

**New file:** `src/NM.SwAddin/Drawing/DrawingDimensioner.cs`

Port dimensioning logic:

| VBA Function | C# Method |
|---|---|
| `DimensionFlat()` | `DimensionFlatPattern()` |
| `DimensionTube()` | `DimensionTube()` |
| `DimensionOther()` | `DimensionFormedView()` |
| `RectProfile()` | `DimensionRectProfile()` |
| `AlignDims.Align()` | `AlignAllDimensions()` |

**Key implementation details:**
- Build EXTSKETCHSEGMENT selection names: `"{segName}@{sketchName}@{componentName}@{viewName}"`
- Dimension placement coordinates use view transform: `position * origVXf[2] / 39.3700787401575` (inches→meters)
- Bend line dimensions get " BL" suffix via `EditDimensionProperties2`
- Created projected views use `swHIDDEN` then `swHIDDEN_GREYED` display modes

### Phase 5: Enhanced DrawingGenerator (NM.SwAddin)

**File:** `src/NM.SwAddin/Drawing/DrawingGenerator.cs` (modify existing)

Add the missing functionality:

1. **Tube laser path** — create unfolded view, dimension both views
2. **Sheet metal dimensioning** — call `DimensionFlat` after dropping view
3. **View auto-scaling** — iterative scale-up loop (max 25 iterations, 5% per step)
4. **Formed view creation** — `CreateUnfoldedViewAt3()` for bent state
5. **Etch mark visibility** — find ProfileFeature sketches, unblank in drawing view
6. **Reference plane hiding** — iterate features, blank RefPlane types
7. **Template path** — use Northern template (`A-SIZE.drwdot`) as default with configurable override

### Phase 6: Drawing Validation (NEW Feature)

**New file:** `src/NM.SwAddin/Drawing/DrawingValidator.cs`

#### a. Dangling Dimension Detection (Color-Based)

Strategy from user's VBA script — dangling dimensions appear in a specific color in SolidWorks:

```csharp
public List<string> FindDanglingDimensions(IDrawingDoc drawDoc)
{
    // Iterate all annotations in all views
    // For each IDisplayDimension:
    //   - Get the IAnnotation
    //   - Check annotation.Color against the dangling color (usually system red/magenta)
    //   - OR check IDisplayDimension.IsDangling (if available in API version)
    //   - Alternatively: check if attached references are valid
    // Return list of dangling dimension names/descriptions
}
```

**SolidWorks dangling detection approaches (in order of reliability):**
1. **Color check** — SolidWorks displays dangling dimensions in a specific color (default: magenta/red, configurable). Read `IAnnotation.Color` and compare against the system dangling color from `UserPreference`.
2. **IDisplayDimension interface** — check if the dimension's references (GetDisplayData → entities) return null
3. **Annotation.IsDangling** — available in newer API versions

#### b. Bend Line Dimension Completeness

```csharp
public List<string> FindUndimensionedBendLines(IDrawingDoc drawDoc, IView flatPatternView)
{
    // 1. Get all bend lines from view: flatPatternView.GetBendLines()
    // 2. Get all dimensions in the view
    // 3. For each bend line, check if any dimension references it
    //    (match by sketch segment name in the dimension's reference entities)
    // 4. Return bend lines that have no associated dimension
}
```

### Phase 7: Integration & Testing

**Modify:** `src/NM.SwAddin/Pipeline/MainRunner.cs` — add drawing generation step after processing

**Modify:** `src/NM.BatchRunner/Program.cs` — add `--create-drawings` flag

**New QA test parts:** Consider adding a simple sheet metal part with bends to `tests/GoldStandard_Inputs/` if not already covered

---

## File Changes Summary

| File | Action | Description |
|------|--------|-------------|
| `docs/vba-reference/DimensionDrawing.bas` | Add | VBA reference for dimensioning |
| `docs/vba-reference/AlignDims.bas` | Add | VBA reference for dimension alignment |
| `src/NM.Core/Drawing/DrawingTypes.cs` | New | Data types (ProcessedElement, BendElement, etc.) |
| `src/NM.SwAddin/Drawing/DrawingGenerator.cs` | Modify | Add full VBA parity (tube, dims, etch, planes) |
| `src/NM.SwAddin/Drawing/ViewGeometryAnalyzer.cs` | New | Edge/bend finding + coordinate transforms |
| `src/NM.SwAddin/Drawing/DrawingDimensioner.cs` | New | Auto-dimensioning (flat, tube, formed views) |
| `src/NM.SwAddin/Drawing/DrawingValidator.cs` | New | Dangling dims + bend line completeness |
| `swcsharpaddin.csproj` | Modify (via sync-csproj) | Include new files |

## Execution Strategy

**Execution: Sequential** — each phase builds on the previous. The geometry analysis must work before dimensioning can reference edges/bends, and dimensioning must work before validation can verify the output.

**Build order:**
1. Data types (no dependencies)
2. ViewGeometryAnalyzer (depends on data types)
3. DrawingDimensioner (depends on geometry analyzer)
4. DrawingGenerator enhancements (depends on dimensioner)
5. DrawingValidator (depends on generator producing output)
6. Integration + testing

## Risk Areas & Gotchas

| Risk | Mitigation |
|------|-----------|
| `GetPolylines7` not available in C# interop | Fall back to `GetVisibleEntities2` + edge traversal |
| View transform coordinates (inches vs meters) | VBA uses `39.3700787401575` (in/m). Track units carefully |
| EXTSKETCHSEGMENT selection name format | Must match exactly: `"{name}@{sketch}@{component}@{view}"` |
| `DimensionProperties2` signature differs across SW versions | Use the full 17-parameter overload documented in VBA |
| Bend lines from higher-level bends (not in flat pattern) | These won't appear in `GetBendLines()` — this is expected, warn but don't fail |
| Drawing template path hardcoded to network share | Make configurable with fallback to SW default template |
| Dangling dimension color is user-configurable | Read from `swUserPreferenceIntegerValue_e.swDanglingDimensionColor` |

## Success Criteria

1. `build-and-test.ps1` → BUILD: SUCCESS
2. No new warnings introduced
3. Drawing creation produces same output as VBA for test parts:
   - Flat pattern view positioned correctly
   - Overall dimensions present
   - Bend-to-bend dimensions with " BL" suffix present
   - Formed view created for parts with bends
   - Etch marks visible in drawing
4. `DrawingValidator` correctly:
   - Identifies dangling dimensions by color
   - Reports bend lines without dimensions
5. New .cs files synced to csproj
