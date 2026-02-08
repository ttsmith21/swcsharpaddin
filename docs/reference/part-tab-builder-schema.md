# Part Custom Property Tab Builder Schema

Reference for the SolidWorks custom property tab builder XML used for **PARTS** (sheet metal and tube).
Property names here are the exact strings stored as SolidWorks custom properties.

## General Tab

| Property         | Control   | Notes                                    |
|------------------|-----------|------------------------------------------|
| `Customer`       | TextBox   | Customer name                            |
| `Print`          | TextBox   | Part/print number                        |
| `Revision`       | TextBox   | Revision letter/number                   |
| `Description`    | TextBox   | Part description                         |
| `rbPartType`     | RadioBtn  | 0=Part, 1=Machined, 2=Purchased         |
| `cbAttachPrint`  | CheckBox  | Attach print to ERP                      |
| `cbAttachCAD`    | CheckBox  | Attach CAD to ERP                        |
| `cbShip`         | CheckBox  | Ship with order                          |

## Machined/Purchased (conditional on rbPartType)

| Property             | Control  | Notes                                  |
|----------------------|----------|----------------------------------------|
| `rbPartTypeSub`      | RadioBtn | Sub-type                               |
| `PurchasedPartNumber`| TextBox  | Purchased part vendor number           |
| `CustPartNumber`     | TextBox  | Customer part number                   |
| `MP_RN`              | TextBox  | Machined/purchased routing note        |

## Outsource

| Property | Control  | Notes                                        |
|----------|----------|----------------------------------------------|
| `OS_WC`  | ComboBox | Outside work center                          |
| `OS_OP`  | TextBox  | Outside operation number                     |
| `OS_RN`  | TextBox  | Outside routing note                         |

## Material

| Property         | Control   | Values / Notes                           |
|------------------|-----------|------------------------------------------|
| `OptiMaterial`   | TextBox   | Material code (resolved)                 |
| `rbMaterialType` | RadioBtn  | **0**=Sheet Metal, **1**=Bar/Pipe/Tube, **2**=Material by Sq.Ft. |

> **IMPORTANT**: `rbMaterialType` is a RadioButton selector (0/1/2), NOT the material name string.
> The material name is entered/resolved through `OptiMaterial`.

## Work Centers

### OP20 — Cutting (Laser/Waterjet/Plasma)

| Property  | Control   | Notes                                      |
|-----------|-----------|--------------------------------------------|
| `OP20`    | ComboBox  | Work center (from `OP20.txt`: F115, F116, F117, F118, F119, F120, N140, N141, N142) |
| `OP20_S`  | TextBox   | Setup time (minutes)                       |
| `OP20_R`  | TextBox   | Run time (minutes)                         |
| `OP20_RN` | TextBox   | Routing note                               |

### F210 — Deburr (OP30)

| Property        | Control   | Notes                                    |
|-----------------|-----------|------------------------------------------|
| `F210`          | CheckBox  | **"1"** = enabled, **"0"** = disabled    |
| `F210_S`        | TextBox   | Setup time (minutes)                     |
| `F210_R`        | TextBox   | Run time (minutes)                       |
| `F210_RN`       | TextBox   | Routing note                             |
| `EdgeClass`     | ComboBox  | Edge classification                      |
| `SurfaceFinish` | ComboBox  | Surface finish requirement               |

### F220 — Tap/Drill (OP35)

| Property  | Control   | Notes                                      |
|-----------|-----------|--------------------------------------------|
| `F220`    | CheckBox  | **"1"** = enabled, **"0"** = disabled      |
| `F220_S`  | TextBox   | Setup time (minutes)                       |
| `F220_R`  | TextBox   | Run time (minutes)                         |
| `F220_RN` | TextBox   | Routing note                               |

### F140 — Press Brake (OP40)

| Property           | Control   | Notes                                 |
|--------------------|-----------|---------------------------------------|
| `PressBrake`       | CheckBox  | **"Checked"** / **"Unchecked"**       |
| `F140_S`           | TextBox   | Setup time (minutes)                  |
| `F140_R`           | TextBox   | Run time (minutes)                    |
| `BrakeLoc`         | TextBox   | Brake location                        |
| `PunchRadius`      | TextBox   | Punch radius                          |
| `VDieWidth`        | TextBox   | V-die width                           |
| `HalfBendDeduction`| TextBox   | Half bend deduction                   |

### F325 — Roll Forming (OP50)

| Property  | Control   | Notes                                      |
|-----------|-----------|--------------------------------------------|
| `F325`    | CheckBox  | **"1"** = enabled, **"0"** = disabled      |
| `F325_S`  | TextBox   | Setup time (minutes)                       |
| `F325_R`  | TextBox   | Run time (minutes)                         |

### Other WC Slots (1-6)

Flexible work center slots for operations not covered by the fixed slots above.
Used for: Weld, Inspect, Hardware, Machine, etc.

**Slot 1** (default OP60):

| Property      | Control   | Notes                                    |
|---------------|-----------|------------------------------------------|
| `OtherWC_CB`  | CheckBox  | Enable slot                              |
| `OtherOP`     | TextBox   | Operation number (default: 60)           |
| `Other_WC`    | ComboBox  | Work center code                         |
| `Other_S`     | TextBox   | Setup time (minutes)                     |
| `Other_R`     | TextBox   | Run time (minutes)                       |
| `Other_RN`    | TextBox   | Routing note                             |

**Slot N (2-6)** (default OP 70, 80, 90, 100, 110):

| Property         | Control   | Notes                                 |
|------------------|-----------|---------------------------------------|
| `OtherWC_CB{N}`  | CheckBox  | Enable slot                           |
| `Other_OP{N}`    | TextBox   | Operation number (default: 60+N*10)   |
| `Other_WC{N}`    | ComboBox  | Work center code                      |
| `Other_S{N}`     | TextBox   | Setup time (minutes)                  |
| `Other_R{N}`     | TextBox   | Run time (minutes)                    |
| `Other_RN{N}`    | TextBox   | Routing note                          |

> **NOTE**: Slot 1 uses `OtherOP` (no underscore), slots 2-6 use `Other_OP{N}` (with underscore).

## Programming Tab

| Property      | Control   | Notes                                      |
|---------------|-----------|--------------------------------------------|
| `Grain`       | CheckBox  | Grain direction                            |
| `ComCut`      | CheckBox  | Common cut                                 |
| `Same`        | CheckBox  | Same as another part                       |
| `External`    | CheckBox  | External programming                       |
| `SheetEdge`   | CheckBox  | Sheet edge                                 |
| `180Degrees`  | CheckBox  | 180-degree rotation                        |
