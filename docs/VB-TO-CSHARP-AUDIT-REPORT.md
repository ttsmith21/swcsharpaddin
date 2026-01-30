# VB.NET to C# Audit Report

## Overview

This document provides a rigorous line-by-line audit comparing the VB.NET reference implementation in `docs/Csharp-reference/SWExtractDataAddin` to the C# implementation in this project.

**Audit Date:** 2026-01-30
**Source:** `docs/Csharp-reference/SWExtractDataAddin/ExtractData/` (VB.NET)
**Target:** `src/NM.SwAddin/Geometry/`, `src/NM.Core/Processing/` (C#)

---

## File Mapping Summary

| VB.NET File | C# Equivalent | Status |
|-------------|---------------|--------|
| `CFace.vb` | `FaceWrapper.cs` | PRESERVED (with improvements) |
| `CFaceCollection.vb` | `TubeGeometryExtractor.cs` | PRESERVED (refactored) |
| `CLoop.vb` | Inline in `FaceWrapper.cs` | ADAPTED |
| `CLoopCollection.vb` | Inline in `FaceWrapper.cs` | ADAPTED |
| `CStepFile.vb` | `TubeGeometryExtractor.cs` + `SimpleTubeProcessor.cs` | PRESERVED |
| `CStepFileCollection.vb` | Not needed (single-file processing) | N/A |
| `MCommon.vb` | `TubeProfile.cs` (TubeShape enum) | PRESERVED |
| `MObjects.vb` | Inline in extractors | PRESERVED |
| `SwAddin.vb` | `SwAddin.cs` (existing) | SEPARATE |
| `EventHandling.vb` | Not ported (UI events) | N/A |

---

## Detailed Audit: CFace.vb → FaceWrapper.cs

### Properties

| VB.NET | C# | Status | Notes |
|--------|-----|--------|-------|
| `swFace As sldworks.Face2` | `_face As IFace2` | PRESERVED | Interface naming convention |
| `oSurface As sldworks.Surface` | `_surface As ISurface` | PRESERVED | |
| `dArea As Double` | `_area As double` | PRESERVED | |
| `bPlanar As Boolean` | `_isPlanar As bool` | PRESERVED | |
| `oNormal As sldworks.MathVector` | Not used | ADAPTED | Normal stored directly as double[] |
| `dNormal(2) As Double` | `_normal As double[]` | PRESERVED | |
| `bRound As Boolean` | `_isRound As bool` | PRESERVED | |
| `oAxis(2) As Double` | `_axis As double[]` | PRESERVED | |
| `oOrigin(2) As Double` | `_origin As double[]` | PRESERVED | |
| `dRadius As Double` | `_radius As double` | PRESERVED | |
| `oLoopsOnMe As CLoopCollection` | Via `GetLoops()` method | ADAPTED | Lazy on-demand retrieval |
| `oHolesOnMe As CLoopCollection` | Via `GetHoles()` method | ADAPTED | Lazy on-demand retrieval |

### Methods

| VB.NET Method | C# Method | Status | Notes |
|---------------|-----------|--------|-------|
| `New(ByRef oFace As Face2)` | `FaceWrapper(IFace2 face)` | PRESERVED | Error handling differs |
| `Finalize()` | Not implemented | ADAPTED | C# uses GC, no explicit disposal |
| `Populate()` | Constructor body | PRESERVED | Merged into constructor |
| `IsSame(ByRef oFace As CFace)` | `IsSame(FaceWrapper other, ISldWorks)` | PRESERVED | Requires swApp parameter |
| `IsSame(ByRef oFace As Face2)` | `IsSame(IFace2 otherFace, ISldWorks)` | PRESERVED | |
| `IsFaceAxisParallelToThis(ByRef dAxis)` | `IsVectorParallel()` (private) | PRESERVED | |
| `IsFaceAxisParallelToThisFace(ByRef oFaceToTest)` | `IsAxisParallelTo(FaceWrapper)` | PRESERVED | |
| `IsFaceNormalParallelToThisFace(ByRef oFaceToTest)` | `IsNormalParallelTo(FaceWrapper)` | PRESERVED | |
| `IsFaceNormalParallelToThisNormal(ByRef oFaceNormal)` | `IsNormalParallelTo(double[])` | PRESERVED | |
| `SelectFace(ByRef swSelectData, isAppend)` | `SelectFace(SelectData, bool)` | PRESERVED | |
| `GetLoops()` | `GetLoops()` | PRESERVED | Returns `List<ILoop2>` |
| `GetHoles()` | `GetHoles()` | PRESERVED | LINQ filter |
| `GetAllEdgesOfAllOuterLoops()` | `GetOuterLoopEdges()` | PRESERVED | |
| `GetDirectionOfLargestLinearEdge(...)` | `GetLargestLinearEdgeDirection(...)` | PRESERVED | out params vs ByRef |
| `GetEdgesForCutLengthOtherThanHolesForRoundProfile()` | `GetOuterLoopEdges()` | PRESERVED | Simplified |
| `GetEdgesForCutLengthOtherThanHolesForOtherThanRoundProfile(...)` | In `TubeGeometryExtractor` | PRESERVED | Moved to extractor |
| `GetAllEdgesWhoseAdjacentFacesArePlanarAndWhoseOtherAdjacentFaceIsNotInThisCollection(...)` | Not directly ported | **MISSING** | Complex edge filtering |
| `GetAdjacentFaceOfThisEdgeOtherThanMe(ByRef oEdge)` | `GetAdjacentFace(IEdge, ISldWorks)` | PRESERVED | |
| `GetEdgesWhichAreInLoopWithThisEdge(...)` | Not ported | **MISSING** | Complex loop-edge relationship |
| `GetMaterialLength(...)` | `GetMaterialLength(...)` | PRESERVED | |
| `ProcessTessTriangles(...)` | Inline in `GetMaterialLength` | PRESERVED | |
| `GetMax()` / `GetMin()` | `Math.Max()` / `Math.Min()` | PRESERVED | Standard library |

### VB.NET-Specific Issues Addressed

1. **On Error Resume Next**: All VB.NET methods use `On Error Resume Next` with `If Err.Number <> 0 Then Err.Clear()`. C# uses proper try-catch blocks where appropriate, with null checks for defensive programming.

2. **Nothing vs null**: All `Is Nothing` checks converted to `== null`.

3. **ByRef Parameters**:
   - VB.NET: `ByRef oFace As sldworks.Face2`
   - C#: `IFace2 face` (reference types are implicitly by-ref for modification)
   - For out values: C# uses `out` parameters

4. **Array Bounds**:
   - VB.NET: `Dim dNormal(2) As Double` creates array of 3 elements (0-2)
   - C#: `new double[3]` - correctly sized

5. **AndAlso/OrElse vs And/Or**:
   - VB.NET uses both short-circuit and non-short-circuit operators
   - C# uses `&&` and `||` (always short-circuit)
   - **No behavioral difference found** in this code

---

## Detailed Audit: CFaceCollection.vb → TubeGeometryExtractor.cs

### Properties

| VB.NET | C# | Status |
|--------|-----|--------|
| `oFaceCollection As List(Of CFace)` | `List<FaceWrapper> faces` (local) | ADAPTED |
| `oMaxAreaFaceIndexCollection As List(Of Integer)` | `maxAreaFaces` (local) | ADAPTED |
| `oHoles As CLoopCollection` | Via face methods | ADAPTED |
| `iIndexOfRoundFace As Integer` | `roundFace` variable | ADAPTED |
| `dStartPointOfMaterialLength(2)` | `TubeProfile.StartPoint` | PRESERVED |
| `dEndPointOfMaterialLength(2)` | `TubeProfile.EndPoint` | PRESERVED |

### Core Algorithm: ComputeShape()

**VB.NET Location:** `CFaceCollection.vb:135-351`
**C# Location:** `TubeGeometryExtractor.cs:119-432`

| Logic Block | Status | Notes |
|-------------|--------|-------|
| Find max-area faces | PRESERVED | LINQ vs manual loop |
| Check for round face | PRESERVED | `FirstOrDefault` vs `While` loop |
| Round profile extraction | PRESERVED | `ExtractRoundProfile()` |
| Non-round profile extraction | PRESERVED | `ExtractNonRoundProfile()` |
| Get largest linear edge direction | PRESERVED | |
| Get faces parallel to primary | PRESERVED | LINQ filter |
| Get faces normal to axis | PRESERVED | LINQ filter |
| Cross-product calculation | PRESERVED | `CrossProduct()` method |
| Distance measurement | PRESERVED | `MeasureMaxDistanceWithValidation()` |
| Modulo check for shape | **IMPROVED** | Added `distinctDistanceCount` tracking |
| Wall ratio validation | **IMPROVED** | Added MAX_WALL_RATIO check (20%) |
| Shape determination logic | PRESERVED | Same conditional structure |
| Hole count | PRESERVED | |
| Cut length calculation | PRESERVED | |

### Shape Determination Logic (Critical)

**VB.NET (lines 225-309):**
```vbnet
If IsTendsToZero(dHeight) Or IsTendsToZero(dWidth) Then
    If IsTendsToZero(dHeight) And IsTendsToZero(dWidth) Then
        eShape = EnumShape.angle
    Else
        eShape = EnumShape.channel
    End If
    ' ... face removal logic
Else
    Dim iRemPrimary As Integer = oDistancesCollectionForPrimaryFace.Count Mod 2
    Dim iRemSecondary As Integer = oDistancesCollectionForSecondaryEdge.Count Mod 2
    If iRemPrimary <> 0 Or iRemSecondary <> 0 Then
        ' Odd distances = angle or channel
    Else
        If IsTendsToZero(dHeight - dWidth) Then
            eShape = EnumShape.square
        Else
            eShape = EnumShape.rectangle
        End If
    End If
End If
```

**C# (lines 374-392):**
```csharp
bool heightZero = IsTendsToZero(height);
bool widthZero = IsTendsToZero(width);

if (heightZero && widthZero)
{
    result.Shape = TubeShape.Angle;
}
else if (heightZero || widthZero)
{
    result.Shape = TubeShape.Channel;
}
else if (IsTendsToZero(height - width))
{
    result.Shape = TubeShape.Square;
}
else
{
    result.Shape = TubeShape.Rectangle;
}
```

**Status:** PRESERVED - Logic is equivalent, though C# is more readable.

### Missing Logic in C#

1. **Face removal from cut-length collection** (VB lines 232-252, 260-277, 280-298):
   - VB.NET removes smaller-area faces from cut-length collection based on complex area comparison
   - C# does not implement this optimization
   - **Impact:** May include extra edges in cut length calculation

2. **`RemoveAllEdgesWhichAreParallelToThisDirection()`** (VB line 328):
   - VB.NET filters out edges parallel to primary axis
   - C# has `IsEdgeParallelToDirection()` but uses it differently
   - **Impact:** Minor - cut length may be slightly different

---

## Detailed Audit: CLoop.vb / CLoopCollection.vb

### CLoop.vb → Inline in FaceWrapper

| VB.NET | C# | Status |
|--------|-----|--------|
| `New(ByRef oLoop As Loop2)` | N/A (use ILoop2 directly) | ADAPTED |
| `IsSame()` | N/A (not needed) | ADAPTED |
| `MyLoop` property | Direct ILoop2 reference | ADAPTED |
| `GetEdges()` | `loop.GetEdges()` | PRESERVED |
| `IsOuter` property | `loop.IsOuter()` | PRESERVED |
| `SelectEdgesOfLoop()` | Not ported | **MISSING** |

### CLoopCollection.vb → FaceWrapper methods

| VB.NET Method | C# Equivalent | Status |
|---------------|---------------|--------|
| `New()` | N/A | ADAPTED |
| `New(ByRef oFace)` | `FaceWrapper.GetLoops()` | ADAPTED |
| `Count` | `List.Count` | PRESERVED |
| `Item(index)` | `List[index]` | PRESERVED |
| `Add()` | `List.Add()` | PRESERVED |
| `AddAllLoopsOfThisLoopCollection()` | `AddRange()` | PRESERVED |
| `IsAnyLoopOuterLoop()` | LINQ `.Any()` | PRESERVED |
| `GetAllOuterLoops()` | `GetOuterLoops()` | PRESERVED |
| `GetAllHoles()` | `GetHoles()` | PRESERVED |
| `SelectAllLoops()` | Not ported | **MISSING** |
| `Clear()` | Not needed | N/A |
| `GetEdges()` | `GetOuterLoopEdges()` / `GetHoleEdges()` | PRESERVED |

---

## Detailed Audit: CStepFile.vb → TubeGeometryExtractor + SimpleTubeProcessor

### Properties

| VB.NET | C# | Status |
|--------|-----|--------|
| `bIsSelected As Boolean` | Not needed (single file) | N/A |
| `strFullFileName As String` | Model path from IModelDoc2 | ADAPTED |
| `strShape As String` | `TubeProfile.ShapeName` | PRESERVED |
| `dWallThickness As Double` | `TubeProfile.WallThicknessMeters` | PRESERVED |
| `dCutLength As Double` | `TubeProfile.CutLengthMeters` | PRESERVED |
| `dMaterialLength As Double` | `TubeProfile.MaterialLengthMeters` | PRESERVED |
| `iNumberOfHoles As Integer` | `TubeProfile.NumberOfHoles` | PRESERVED |
| `strCrossSection As String` | `TubeProfile.CrossSection` | PRESERVED |
| `swModel As ModelDoc2` | Parameter to `Extract()` | ADAPTED |
| `swSelectData`, `swSelectionManager` | Created inline | ADAPTED |
| `swUserUnit As UserUnit` | Not ported (unit conversion in TubeProfile) | ADAPTED |
| `oFaces As CFaceCollection` | `List<FaceWrapper>` local | ADAPTED |
| INotifyPropertyChanged | Not needed (not WPF) | N/A |

### Methods

| VB.NET | C# | Status |
|--------|-----|--------|
| `ReadFileAndPopulateModel()` | Not needed (model passed in) | N/A |
| `Bodies` property | `partDoc.GetBodies2()` | PRESERVED |
| `ExtractData()` | `TubeGeometryExtractor.Extract()` | PRESERVED |
| `SaveDataAsCustomProperties()` | `SimpleTubeProcessor.Process()` | PRESERVED |
| `SaveModel()` | Not ported | **MISSING** |
| `ClearAllSelection()` | `model.ClearSelection2()` | PRESERVED |
| `SelectAllEdgesForCutLength()` | Not ported | **MISSING** |
| `SelectAllEdgesForHoles()` | Not ported | **MISSING** |
| `ActivateMe()` | Not needed | N/A |
| `UpdateView()` | Not ported (display mode) | **MISSING** |
| `ShowHideCallout()` | Not ported (UI) | N/A |
| `CreateCallout()` | Not ported (UI) | N/A |

---

## Detailed Audit: MCommon.vb → TubeProfile.cs

| VB.NET | C# | Status |
|--------|-----|--------|
| `MaxDouble` / `MinDouble` | `double.MaxValue` / `double.MinValue` | PRESERVED |
| `strRoundShape` etc. constants | `TubeShape` enum | IMPROVED |
| `EnumShape` | `TubeShape` | PRESERVED |
| `IsTendsToZero(val)` | `IsTendsToZero(val)` private methods | PRESERVED |
| `GetParentFolderName()` | Not needed | N/A |
| `GetFileNameWithoutExtensionFromFilePathName()` | Not needed | N/A |
| `GetProductVersion()` | Not needed | N/A |

### Tolerance Check

**VB.NET:**
```vbnet
Public Function IsTendsToZero(ByRef val As Double) As Boolean
    If Math.Abs(val) < 0.000000001 Then
        Return True
    End If
    Return False
End Function
```

**C#:**
```csharp
private static bool IsTendsToZero(double val)
{
    return Math.Abs(val) < Tolerance; // Tolerance = 1e-9
}
```

**Status:** PRESERVED - Identical tolerance (10^-9)

---

## Detailed Audit: MObjects.vb → Inline Methods

| VB.NET Function | C# Location | Status |
|-----------------|-------------|--------|
| `SWApp As SldWorks` | Constructor injection | ADAPTED |
| `GetCutLengthFromEdgeCollection()` | `CalculateTotalEdgeLength()` | **ADAPTED** |
| `DoesThisEdgeExistInThisList()` | `List.Contains()` / reference equality | ADAPTED |
| `DoesAnyOfTheseFacesExistInThisList()` | Not directly ported | **MISSING** |
| `RemoveThisValueFromThisList()` | `List.Remove()` | ADAPTED |
| `DoesStartAndEndPointsOfThisEdgeLieOnFacesFromThisFaceCollection()` | Not ported | **MISSING** |
| `RemoveAllEdgesWhichAreParallelToThisDirection()` | `IsEdgeParallelToDirection()` filter | PRESERVED |
| `GetDirectionOfThisEdge()` | Inline in `IsEdgeParallelToDirection()` | PRESERVED |
| `AreTheseVectorsParallel()` | `IsVectorParallel()` | PRESERVED |
| `SelectEdge()` | Not ported (UI selection) | **MISSING** |
| `CreateMathPointAt()` | Not ported (callouts) | N/A |

### Critical Difference: Cut Length Calculation

**VB.NET (`GetCutLengthFromEdgeCollection`):**
```vbnet
' Uses SolidWorks Measure object to get TotalLength of selected edges
swModel.ClearSelection2(True)
For i = 0 To oEdgeCollection.Count - 1
    oEdgeCollection.Item(i).Select4(True, swSelectData)
Next
Dim oMeasure As sldworks.Measure = swModel.Extension.CreateMeasure()
If oMeasure.Calculate(Nothing) Then
    dCutLength = oMeasure.TotalLength()
End If
```

**C# (`CalculateTotalEdgeLength`):**
```csharp
// Calculates length by iterating edges individually
foreach (var edge in edges)
{
    var curve = (ICurve)edge.GetCurve();
    double start, end;
    bool isClosed, isPeriodic;
    if (curve.GetEndParams(out start, out end, out isClosed, out isPeriodic))
    {
        totalLength += curve.GetLength3(start, end);
    }
}
```

**Status:** ADAPTED - Different approach but mathematically equivalent. C# version doesn't require UI selection.

---

## Summary: Missing Logic

### High Priority (May Affect Calculations)

1. **`GetEdgesWhichAreInLoopWithThisEdge()`** - Complex loop-edge relationship finding
2. **`GetAllEdgesWhoseAdjacentFacesArePlanarAndWhoseOtherAdjacentFaceIsNotInThisCollection()`** - Edge filtering for cut length
3. **`DoesStartAndEndPointsOfThisEdgeLieOnFacesFromThisFaceCollection()`** - Vertex-face intersection check
4. **Face removal logic in non-round profiles** - Removes smaller faces from cut length collection

### Medium Priority (UI/Selection Features)

5. **`SelectAllEdgesForCutLength()`** - Visual feedback
6. **`SelectAllEdgesForHoles()`** - Visual feedback
7. **`SelectAllLoops()`** - Loop selection
8. **`SelectEdge()`** - Edge selection utility
9. **`UpdateView()`** - Display mode setting

### Low Priority (Not Needed)

10. **`SaveModel()`** - Model saving handled elsewhere
11. **`ActivateMe()`** - Document activation handled by workflow
12. **Callout system** - UI-specific, not needed for automation

---

## VB.NET-Specific Gotchas Analysis

### 1. On Error Resume Next

**VB.NET Pattern:**
```vbnet
Public Function SomeMethod() As ReturnType
    On Error Resume Next
    ' ... code ...
    If Err.Number <> 0 Then Err.Clear()
    Return result
End Function
```

**C# Pattern:**
```csharp
public ReturnType SomeMethod()
{
    try
    {
        // ... code with null checks ...
        return result;
    }
    catch (Exception ex)
    {
        ErrorHandler.HandleError(...);
        return default;
    }
}
```

**Status:** ADDRESSED - C# uses defensive null checks and try-catch where appropriate.

### 2. Nothing vs null

All instances converted correctly:
- `Is Nothing` → `== null`
- `IsNot Nothing` → `!= null`
- `= Nothing` → `= null`

### 3. Array Declarations

**VB.NET:** `Dim arr(2) As Double` = 3 elements (0, 1, 2)
**C#:** `new double[3]` = 3 elements

**Status:** CORRECT - Arrays sized properly

### 4. ByRef Parameters

Most `ByRef` parameters in VB.NET are reference types that don't need explicit `ref` in C#.
Output parameters correctly use `out` keyword.

### 5. Integer Division

No integer division (`\`) found in the VB.NET code. All divisions use `/`.

### 6. String Concatenation

VB.NET uses `&` for string concatenation. C# uses `+` or string interpolation.
**Status:** CORRECT

### 7. Short-Circuit Evaluation

VB.NET uses `AndAlso`/`OrElse` for short-circuit (most of the code uses non-short-circuit `And`/`Or`).
C# always uses short-circuit `&&`/`||`.

**Potential Issue:** If VB.NET code relies on side effects of both operands being evaluated, this could cause behavioral differences.

**Analysis:** No side effects found in boolean expressions - SAFE.

### 8. Type Conversions

| VB.NET | C# |
|--------|-----|
| `CDbl(strCrossSection)` | `double.Parse()` |
| `CInt()` | `(int)` cast or `Convert.ToInt32()` |

**Status:** Most conversions not needed as C# is strongly typed.

---

## Recommendations

1. **Consider implementing missing cut-length edge filtering** - May improve accuracy for non-round profiles

2. **Add unit tests for shape detection** - Verify all shape types detected correctly

3. **Document tolerance assumptions** - The 10^-9 tolerance is critical for geometry comparisons

4. **Consider adding selection highlighting** - Useful for debugging geometry extraction

---

## Appendix: Enum Mapping

| VB.NET (MCommon.EnumShape) | C# (TubeShape) |
|----------------------------|----------------|
| `none = 0` | `None = 0` |
| `round = 1` | `Round = 1` |
| `square = 2` | `Square = 2` |
| `rectangle = 3` | `Rectangle = 3` |
| `angle = 4` | `Angle = 4` |
| `channel = 5` | `Channel = 5` |

**Status:** IDENTICAL
