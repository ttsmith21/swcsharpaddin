---
name: sw-api-patterns
description: SolidWorks API patterns for C#. Use when writing SW API code.
---

# SolidWorks API Patterns

## Getting Active Document
```csharp
var swApp = (ISldWorks)Application;
var model = swApp.ActiveDoc as IModelDoc2;
if (model == null) return; // Always null-check
```

## Traversing Features
```csharp
var feat = model.FirstFeature() as IFeature;
while (feat != null)
{
    // Process feature...
    feat = feat.GetNextFeature() as IFeature;
}
```

## Safe Array Handling
```csharp
public static List<T> SafeCast<T>(object swArray) where T : class
{
    if (swArray == null) return new List<T>();
    return ((object[])swArray).Cast<T>().ToList();
}

// Usage:
var bodies = SafeCast<IBody2>(swPart.GetBodies2((int)swBodyType_e.swSolidBody, true));
```

## Getting Bodies from Part
```csharp
var swPart = model as IPartDoc;
if (swPart == null) return;

var bodiesRaw = swPart.GetBodies2((int)swBodyType_e.swSolidBody, true);
if (bodiesRaw == null) return;

var bodies = ((object[])bodiesRaw).Cast<IBody2>().ToList();
foreach (var body in bodies)
{
    // Process body...
}
```

## Getting Faces from Body
```csharp
var facesRaw = body.GetFaces();
if (facesRaw == null) continue;

var faces = ((object[])facesRaw).Cast<IFace2>().ToList();
foreach (var face in faces)
{
    var surface = face.GetSurface() as ISurface;
    if (surface == null) continue;

    if (surface.IsCylinder())
    {
        // Handle cylindrical face
    }
    else if (surface.IsPlane())
    {
        // Handle planar face
    }
}
```

## Reading Custom Properties
```csharp
var ext = model.Extension;
var mgr = ext.CustomPropertyManager[""];  // "" = active config, or config name

string valOut, resolvedValOut;
bool wasResolved;
int result = mgr.Get5("PropertyName", false, out valOut, out resolvedValOut, out wasResolved);

if (result == (int)swCustomInfoGetResult_e.swCustomInfoGetResult_ResolvedValue)
{
    // Use resolvedValOut
}
```

## Writing Custom Properties
```csharp
var mgr = model.Extension.CustomPropertyManager[""];

// Add or update
int result = mgr.Add3(
    "PropertyName",
    (int)swCustomInfoType_e.swCustomInfoText,
    "PropertyValue",
    (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd
);
```

## Assembly Traversal
```csharp
var swAssy = model as IAssemblyDoc;
if (swAssy == null) return;

var compsRaw = swAssy.GetComponents(true); // true = top-level only
if (compsRaw == null) return;

var components = ((object[])compsRaw).Cast<IComponent2>().ToList();
foreach (var comp in components)
{
    if (comp.IsSuppressed()) continue;

    var compModel = comp.GetModelDoc2() as IModelDoc2;
    if (compModel == null) continue;

    // Process component...
}
```

## Selection Manager Pattern
```csharp
var selMgr = model.SelectionManager as ISelectionMgr;
int count = selMgr.GetSelectedObjectCount2(-1);

for (int i = 1; i <= count; i++)  // 1-based!
{
    var selType = (swSelectType_e)selMgr.GetSelectedObjectType3(i, -1);
    var obj = selMgr.GetSelectedObject6(i, -1);

    if (selType == swSelectType_e.swSelFACES)
    {
        var face = obj as IFace2;
        // Process face...
    }
}
```

## Selecting Entities Programmatically
```csharp
// Select by name (features, planes, etc.)
bool success = model.Extension.SelectByID2(
    "Top",                              // Name
    "PLANE",                            // Type string
    0, 0, 0,                            // X, Y, Z (for point selection)
    false,                              // Append to selection
    0,                                  // Mark
    null,                               // Callout
    (int)swSelectOption_e.swSelectOptionDefault
);

// Select entity directly
var selData = selMgr.CreateSelectData();
selData.Mark = 1;
bool selected = ((IEntity)face).Select4(true, selData);
```

## Bounding Box
```csharp
var box = (double[])body.GetBodyBox();
if (box != null && box.Length == 6)
{
    double xMin = box[0], yMin = box[1], zMin = box[2];
    double xMax = box[3], yMax = box[4], zMax = box[5];

    double length = xMax - xMin;
    double width = yMax - yMin;
    double height = zMax - zMin;
}
```

## Math Transform (Component Position)
```csharp
var transform = comp.Transform2 as IMathTransform;
if (transform != null)
{
    var arrayData = (double[])transform.ArrayData;
    // arrayData contains 16 elements: 3x3 rotation + translation + scale
    double x = arrayData[9];
    double y = arrayData[10];
    double z = arrayData[11];
}
```

## Error Handling Pattern
```csharp
public void ProcessPart(IModelDoc2 model)
{
    ErrorHandler.PushCallStack("ProcessPart");
    try
    {
        // Your code here...
    }
    catch (Exception ex)
    {
        ErrorHandler.HandleError("ProcessPart", ex.Message, ex);
    }
    finally
    {
        ErrorHandler.PopCallStack();
    }
}
```

## Common Gotchas

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
| Interface casting | Use `as` keyword: `model as IPartDoc` (returns null if wrong type) |
| COM lifetime | Don't call `Marshal.ReleaseComObject` - let COM handle it |
| Threading | All SW API calls must be on main STA thread |
