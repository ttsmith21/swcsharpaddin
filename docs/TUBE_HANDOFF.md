# Tube Processing - Handoff Documentation

## Status: STUBBED - Needs Geometry Extraction Implementation

### What's Complete ?
All tube processing **services** are implemented and ready to use:

1. **PipeScheduleService** - Resolves OD/wall thickness from pipe schedules
2. **TubeMaterialCodeGenerator** - Maps material types to codes (A36?BLK, ALNZD?HR, etc.)
3. **TubeCuttingParameterService** - Provides feed rates and pierce times for laser cutting
4. **RoundBarValidator** - Validates standard OptiMaterial round bar sizes
5. **ExternalStartAdapter** - Optional integration for external automation (16"×0.5" rule)
6. **Work Center Operations** - F325/F140/F210 threshold logic and precision rounding

### What's Blocked ??
**SimpleTubeProcessor** cannot identify or process tube parts because:

**PRIMARY ISSUE:** Tube geometry extraction from SolidWorks is not implemented.

The processor needs to:
1. Detect if a part is a tube (cylindrical body)
2. Extract tube dimensions (OD, ID, length, axis orientation)
3. Pass that data to the ready services above

### The Problem
File: `src\NM.Core\Processing\SimpleTubeProcessor.cs`

#### Current Behavior
- `CanProcess(IModelDoc2 model)` ? Always returns `false` (stubbed)
- `Process(IModelDoc2 model, ProcessingOptions options)` ? Returns `true` without doing anything

#### What Needs Implementation
Method: `ExtractTubeGeometry(IModelDoc2 model)`

**Required Steps:**
```csharp
1. Get the solid body from the part
   var part = model as IPartDoc;
   var bodies = part.GetBodies2((int)swBodyType_e.swSolidBody, false) as object[];
   
2. Iterate through faces looking for cylindrical surfaces
   var body = bodies?[0] as IBody2;
   var faces = body.GetFaces() as object[];
   foreach (IFace2 face in faces)
   {
       var surf = face.IGetSurface();
       if (surf.IsCylinder())
       {
           // Found a cylindrical face
       }
   }

3. From cylindrical face, extract geometry
   var cyl = surf.GetCylinderParams2() as double[];
   // cyl[0..2] = origin point [x, y, z]
   // cyl[3..5] = axis direction [x, y, z]  
   // cyl[6] = radius

4. Identify end caps (circular planar faces)
   - Should have exactly 2 circular planar faces
   - Calculate tube length from distance between centers

5. Determine wall thickness
   Option A: Check for sheet metal feature (if tube was made from rolled sheet)
   Option B: Find inner cylindrical face, compare radii: (OD - ID) / 2

6. Return TubeGeometry data structure
   return new TubeGeometry
   {
       OuterDiameter = radius * 2 * 39.3701, // meters to inches
       WallThickness = ...,
       Length = ...,
       Axis = new[] { cyl[3], cyl[4], cyl[5] }
   };
```

### Debugging Tips

#### Common SolidWorks API Pitfalls
- **Units:** SolidWorks internal units are **meters**. Convert to inches for display/calculations.
- **Face Iteration:** Use `IBody2.GetFaces()` and cast to `object[]`, then iterate as `IFace2`
- **Surface Types:** Use `ISurface.IsCylinder()`, `ISurface.IsPlanar()`, etc. to identify face types
- **Cylinder Parameters:** `GetCylinderParams2()` returns a `double[]` with 7 elements
- **Threading:** All SolidWorks COM calls must be on the **main thread** (STA)

#### Validation Checklist
A valid tube should have:
- ? Exactly 1 solid body (no multi-body parts)
- ? Exactly 2 circular planar faces (end caps)
- ? At least 1 cylindrical face (outer wall)
- ? Optionally 1 inner cylindrical face (for hollow tubes)
- ? Consistent axis direction across all cylindrical faces

#### Test Cases to Try
1. **Simple Hollow Tube** - Cylinder with constant OD/ID
2. **Rolled Sheet Metal Tube** - Has sheet metal feature
3. **Solid Rod** - Single cylindrical face, no inner diameter
4. **Tapered Tube** - Should reject (not a simple tube)
5. **Multi-body Part** - Should reject in validation phase

### Integration Points

Once `ExtractTubeGeometry()` works, the processor will:

1. **Validate** the part is a tube via `CanProcess()`
2. **Extract** geometry (OD, ID, length, axis)
3. **Resolve** schedule using `PipeScheduleService`
4. **Generate** material code using `TubeMaterialCodeGenerator`
5. **Calculate** cutting parameters using `TubeCuttingParameterService`
6. **Determine** work center operations (F325, F140, F210)
7. **Write** custom properties to the model
8. **Invoke** `ExternalStartAdapter` if applicable (16"×0.5" rule)

All the downstream services are **ready and tested** - they just need tube geometry as input.

### Next Steps

1. **Implement** `ExtractTubeGeometry()` in `SimpleTubeProcessor.cs`
2. **Update** `CanProcess()` to call geometry extraction and return true for valid tubes
3. **Update** `Process()` to:
   - Call `ExtractTubeGeometry()`
   - Pass geometry to ready services
   - Write custom properties
   - Handle errors gracefully
4. **Test** with real tube parts in SolidWorks
5. **Validate** property values match VBA macro output

### File References

| File | Status | Purpose |
|------|--------|---------|
| `SimpleTubeProcessor.cs` | ?? Stubbed | Main tube processor - **needs geometry extraction** |
| `PipeScheduleService.cs` | ? Complete | Resolves OD/wall from schedule number |
| `TubeMaterialCodeGenerator.cs` | ? Complete | Material type ? code mapping |
| `TubeCuttingParameterService.cs` | ? Complete | Feed rates and pierce times |
| `RoundBarValidator.cs` | ? Complete | Standard size validation |
| `ExternalStartAdapter.cs` | ? Complete | External automation integration |

### Questions?

The original VBA macro likely has tube geometry extraction code. Look for:
- Face iteration loops
- Cylindrical surface detection
- Radius/diameter calculations
- Length measurements between faces

Port that logic to C# using the SolidWorks Interop API.

**Good luck! The hard infrastructure work is done - just need to crack this geometry nut.**
