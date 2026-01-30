# Gold Standard Corpus Specification

This document defines the "Exhaustive List" of test parts required to validate the SolidWorks Automation Add-in.

**Goal:** Ensure 100% coverage of decision logic (Validation → Classification → Processing).

---

## Class A: Invalid / Problem Parts (The "Reject List")

**Expectation:** Validator catches these immediately. `Status = "Failed"`.

| ID | Name | Description | File Name Convention |
|----|------|-------------|---------------------|
| A1 | Empty File | A file with no geometry at all (0kb logic check) | `A1_EmptyFile.sldprt` |
| A2 | Surface Bodies Only | Imported IGES/STEP that came in as surfaces, not solids | `A2_SurfaceBodiesOnly.sldprt` |
| A3 | Multi-Body Solid | A part file containing separate disjoint bodies (e.g., a weldment saved as a part) | `A3_MultiBody.sldprt` |
| A4 | Suppressed Body | The main body exists but is suppressed in the feature tree | `A4_SuppressedBody.sldprt` |
| A5 | No Material | A solid body with no material assigned | `A5_NoMaterial.sldprt` |
| A6 | Zero Thickness | Geometry that fails geometric check (knife edge) | `A6_ZeroThickness.sldprt` |

---

## Class B: Sheet Metal (The "Happy Path")

**Expectation:** `Classification = SheetMetal`, `Status = Success`.

| ID | Name | Description | Tests | File Name Convention |
|----|------|-------------|-------|---------------------|
| B1 | Standard Bracket (Native) | A simple L-bracket created with SolidWorks Sheet Metal features | Feature detection | `B1_NativeBracket_14ga_CS.sldprt` |
| B2 | Standard Bracket (Imported) | The same L-bracket, but imported as a "dumb" STEP file | ConvertToSheetMetal | `B2_ImportedBracket_14ga_CS.sldprt` |
| B3 | Rolled Cylinder | A rolled plate (open seam) that can be flattened | Unfold cylinder | `B3_RolledCylinder_16ga_SS.sldprt` |
| B4 | Complex Form | A multi-bend part (U-shape with return flanges) | Multi-bend processing | `B4_ComplexUBracket_11ga_CS.sldprt` |
| B5 | With Holes | Part with internal cutouts | Loop validation | `B5_BracketWithHoles_14ga_CS.sldprt` |
| B6 | Tiny Flange | A part with a flange smaller than material thickness | Unfold validation | `B6_TinyFlange_10ga_CS.sldprt` |

---

## Class C: Tube & Structural

**Expectation:** Round Tube = `Tube`, Others = `Generic` (or `SheetMetal` if unfoldable).

| ID | Name | Description | Tests | File Name Convention |
|----|------|-------------|-------|---------------------|
| C1 | Round Tube | Standard hollow cylinder | Concentricity, Wall Thickness | `C1_RoundTube_2OD_SCH40.sldprt` |
| C2 | Rectangular Tube | Hollow box section | Fallback classification | `C2_RectTube_2x1.sldprt` |
| C3 | Square Tube | Equal side box section | Fallback classification | `C3_SquareTube_2x2.sldprt` |
| C4 | Angle Iron | L-profile structural member | Fallback classification | `C4_AngleIron_2x2.sldprt` |
| C5 | C-Channel | Standard C-profile | Fallback classification | `C5_CChannel.sldprt` |
| C6 | I-Beam / H-Beam | Structural beam profile | Fallback classification | `C6_IBeam.sldprt` |
| C7 | Round Bar | Solid cylinder | Should FAIL Tube (no ID), or "Bar" | `C7_RoundBar_1dia.sldprt` |

---

## Class D: Non-Convertible (The "Other" Bucket)

**Expectation:** `Classification = Generic`, `Status = Success` (processed as generic).

| ID | Name | Description | File Name Convention |
|----|------|-------------|---------------------|
| D1 | Machined Block | A milled part with a counter bore and a pockets (cannot unfold) | `D1_MachinedBlock.sldprt` |
| D2 | Elbow Fitting | A pipe elbow (curved axis, not straight tube) | `D2_ElbowFitting.sldprt` |
| D3 | Motor/Gearbox | A purchase component (complex geometry) | `D3_Motor.sldprt` |
| D4 | Fastener | A screw or bolt (threads, hex head) | `D4_Fastener.sldprt` |
| D5 | Casting | Irregular geometry with variable wall thickness | `D5_Casting.sldprt` |
| D6 | U-Bolt | Curved or formed sold rod | `D6_Ubolt.sldprt` |
| D7 | Nut | Nut w/o internal threads hex head | `D7_Nut_No_Threads.sldprt` |
| D8 | Swagelock T-Fitting | T fitting with hex and threads | `D8_Swagelock_T_Fitting.sldprt` |
| D9 | Fastener | A screw or bolt w/o threads (no threads, hex head) | `D9_Fastener_No_Threads.sldprt` |
| D10 | Reducer | A pipe reducer | `D10_Pipe_Reducer.sldprt` |
| D11 | Valve | A pipe Valve | `D11_Pipe_Valve.sldprt` |
| D12 | Stud | A threded stud | `D12_Stud_Threaded.sldprt` |

---

## Class E: Edge Cases (The "Stress Test")

**Expectation:** Specific behavior required.

| ID | Name | Description | Tests | Expected | File Name Convention |
|----|------|-------------|-------|----------|---------------------|
| E1 | Thick Plate | 1" thick plate | MaxSheetThickness | `SheetMetal` or `Generic`? | `E1_ThickPlate_1in.sldprt` |
| E2 | Foil/Shim | 0.005" shim | MinSheetThickness | `SheetMetal` | `E2_Foil_005.sldprt` |
| E3 | Tapered Wall | Sheet metal with draft angle | Top sheet area does not equal bottom sheet area | `Generic` | `E3_TaperedWall.sldprt` |
| E4 | Concentricity Fail | Tube with off-center ID (eccentric) | Concentricity check | `Generic` or `Failed` | `E4_EccentricTube.sldprt` |
| E5 | Non-Planar Face | Bent part with no flat reference face | Unfolding | `Failed` or special handling | `E5_NonPlanarFace.sldprt` |
| E6 | Huge Dimensions | 100-meter long part | BBox overflow/units | Handled gracefully | `E6_HugePart.sldprt` |

---

## Class F: Assemblies (The "Structure Test")

**Expectation:** Correctly traverse BOM and process components.

| ID | Name | Description | Tests | File Name Convention |
|----|------|-------------|-------|---------------------|
| F1 | Single Level Assembly | Simple assembly with 3 unique parts | Basic traversal | `F1_SimpleAssy.sldasm` |
| F2 | Multi-Level Assembly | Assembly containing sub-assemblies | Recursion depth | `F2_MultiLevelAssy.sldasm` |
| F3 | Flexible Sub-Assembly | F1_SimpleAssy Sub-assembly set to "Flexible" | Geometry validation | `F3_FlexibleAssy.sldasm` |
| F4 | Virtual Components | Parts saved internally to assembly | Path resolution | `F4_VirtualComps.sldasm` |
| F5 | Suppressed Components | F1_SimpleAssy and C3_SquareTube_2x2 Components suppressed in active config | Should be skipped | `F5_SuppressedComps.sldasm` |
| F6 | Lightweight Components | Components loaded lightweight | ResolveAllLightweight | `F6_LightweightComps.sldasm` |
| F7 | Envelope Components | B1_NativeBracket_14ga_CS is Enveloped - Reference-only components | Should be excluded | `F7_EnvelopeComps.sldasm` |
| F8 | Missing Reference | Assembly with missing part file | Error handling | `F8_MissingRef.sldasm` |

---

## Class G: Configurations (The "Config Test")

**Expectation:** Process the specific active configuration.

| ID | Name | Description | Tests | File Name Convention |
|----|------|-------------|-------|---------------------|
| G1 | Multi-Config Part | Part with "Long" and "Short" configs | Only process active | `G1_MultiConfig.sldprt` |
| G2 | Sheet Metal Configs | Config A = 10ga, Config B = 14ga | Material/Gauge update | `G2_GaugeConfigs.sldprt` |
| G3 | Flat Pattern Config | Part with derived "SM-FLAT-PATTERN" config | Ignore derived | `G3_FlatPatternConfig.sldprt` |
| G4 | SpeedPak | Assembly with SpeedPak active | Graphics-only skip | `G4_SpeedPak.sldasm` |

---

## Class H: File Formats (The "Import Pipeline")

**Expectation:** Import → Classification → Conversion.

*Note: Most inputs in Classes A-G should be native .sldprt files. Class H is specifically for raw import formats.*

| ID | Name | Description | Tests | File Name Convention |
|----|------|-------------|-------|---------------------|
| H1 | STEP AP203/214 | Standard exchange format | Load → Heal → Detect | `H1_StepImport.step` |
| H2 | IGES | Surfaces or Solid | Gap healing | `H2_IgesImport.igs` |
| H3 | Parasolid (*.x_t) | Native kernel format | Best geometry integrity | `H3_ParasolidImport.x_t` |
| H4 | SAT (ACIS) | Another common kernel format | Load → Classify | `H4_SatImport.sat` |
| H5 | DXF Import | 2D sketch imported as part | Empty/Surface handling | `H5_DxfImport.dxf` |

---

## Coverage Matrix

| Class | Count | Priority | Status |
|-------|-------|----------|--------|
| A - Invalid/Problem | 6 | High | ⬜ Not Started |
| B - Sheet Metal | 6 | High | ⬜ Not Started |
| C - Tube & Structural | 7 | High | ⬜ Not Started |
| D - Non-Convertible | 5 | Medium | ⬜ Not Started |
| E - Edge Cases | 6 | Medium | ⬜ Not Started |
| F - Assemblies | 8 | Medium | ⬜ Not Started |
| G - Configurations | 4 | Low | ⬜ Not Started |
| H - File Formats | 5 | Low | ⬜ Not Started |
| **Total** | **47** | | |

---

## Collection Checklist

Use this checklist to track file collection:

```
[ ] Class A: Invalid Parts (6 files)
    [ ] A1_EmptyFile.sldprt
    [ ] A2_SurfaceBodiesOnly.sldprt
    [ ] A3_MultiBody.sldprt
    [ ] A4_SuppressedBody.sldprt
    [ ] A5_NoMaterial.sldprt
    [ ] A6_ZeroThickness.sldprt

[ ] Class B: Sheet Metal (6 files)
    [ ] B1_NativeBracket_14ga_CS.sldprt
    [ ] B2_ImportedBracket_14ga_CS.sldprt
    [ ] B3_RolledCylinder_16ga_SS.sldprt
    [ ] B4_ComplexUBracket_11ga_CS.sldprt
    [ ] B5_BracketWithHoles_14ga_CS.sldprt
    [ ] B6_TinyFlange_10ga_CS.sldprt

[ ] Class C: Tube & Structural (7 files)
    [ ] C1_RoundTube_2OD_SCH40.sldprt
    [ ] C2_RectTube_2x1.sldprt
    [ ] C3_SquareTube_2x2.sldprt
    [ ] C4_AngleIron_2x2.sldprt
    [ ] C5_CChannel.sldprt
    [ ] C6_IBeam.sldprt
    [ ] C7_RoundBar_1dia.sldprt

[ ] Class D: Non-Convertible (5 files)
    [ ] D1_MachinedBlock.sldprt
    [ ] D2_ElbowFitting.sldprt
    [ ] D3_Motor.sldprt
    [ ] D4_Fastener.sldprt
    [ ] D5_Casting.sldprt

[ ] Class E: Edge Cases (6 files)
    [ ] E1_ThickPlate_1in.sldprt
    [ ] E2_Foil_005.sldprt
    [ ] E3_TaperedWall.sldprt
    [ ] E4_EccentricTube.sldprt
    [ ] E5_NonPlanarFace.sldprt
    [ ] E6_HugePart.sldprt

[ ] Class F: Assemblies (8 files)
    [ ] F1_SimpleAssy.sldasm
    [ ] F2_MultiLevelAssy.sldasm
    [ ] F3_FlexibleAssy.sldasm
    [ ] F4_VirtualComps.sldasm
    [ ] F5_SuppressedComps.sldasm
    [ ] F6_LightweightComps.sldasm
    [ ] F7_EnvelopeComps.sldasm
    [ ] F8_MissingRef.sldasm

[ ] Class G: Configurations (4 files)
    [ ] G1_MultiConfig.sldprt
    [ ] G2_GaugeConfigs.sldprt
    [ ] G3_FlatPatternConfig.sldprt
    [ ] G4_SpeedPak.sldasm

[ ] Class H: File Formats (5 files)
    [ ] H1_StepImport.step
    [ ] H2_IgesImport.igs
    [ ] H3_ParasolidImport.x_t
    [ ] H4_SatImport.sat
    [ ] H5_DxfImport.dxf
```
