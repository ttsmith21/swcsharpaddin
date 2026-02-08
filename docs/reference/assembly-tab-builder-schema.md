# Assembly Custom Property Tab Builder Schema

Reference for the SolidWorks custom property tab builder XML used for **ASSEMBLIES**.
Property names here are the exact strings stored as SolidWorks custom properties.

## General Tab

| Property       | Control   | Notes                                    |
|----------------|-----------|------------------------------------------|
| `Customer`     | TextBox   | Customer name                            |
| `Print`        | TextBox   | Assembly print number                    |
| `Revision`     | TextBox   | Revision letter/number                   |
| `Description`  | TextBox   | Assembly description                     |
| `AttachPrint`  | CheckBox  | Attach print to ERP                      |
| `AttachCAD`    | CheckBox  | Attach CAD to ERP                        |
| `ReqOS_A`      | CheckBox  | Requires outside service                 |
| `cbShip`       | CheckBox  | Ship with order                          |

## Outsource

| Property   | Control  | Notes                                       |
|------------|----------|---------------------------------------------|
| `OS_WC_A`  | ComboBox | Outside work center (assembly)              |
| `OS_OP_A`  | TextBox  | Outside operation number                    |
| `OS_RN_A`  | TextBox  | Outside routing note                        |

## Material

| Property         | Control   | Notes                                    |
|------------------|-----------|------------------------------------------|
| `OptiMaterial`   | TextBox   | Material code                            |
| `SupervisorNote` | TextBox   | Supervisor note                          |
| `DisableUpdate`  | CheckBox  | Disable auto-update                      |

## Routing Operations (OP20 - OP150)

Assembly routing is free-form: each slot can be any operation type.
Slots increment by 10: OP20, OP30, OP40, ... OP150 (14 slots).

> **NOTE**: OP10 is typically reserved for KIT (grouping pieces together).

### Per-Slot Properties

For each slot **OP##** where ## = {20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150}:

| Property    | Control   | Notes                                       |
|-------------|-----------|---------------------------------------------|
| `OP##`      | ComboBox  | Work center (from `WeldOps.txt`)            |
| `OP##_S`    | TextBox   | Setup time (minutes)                        |
| `OP##_R`    | TextBox   | Run time (minutes)                          |
| `OP##_RN`   | TextBox   | Routing note                                |

> **IMPORTANT**: The work center is stored in the `OP##` property itself (e.g., `OP20`, `OP30`),
> NOT `OP##_WC`. This differs from the pattern one might expect.

### Example

| Property  | Value           | Meaning                              |
|-----------|-----------------|--------------------------------------|
| `OP20`    | `F400`          | Work center = Weld (OP20)            |
| `OP20_S`  | `15`            | 15 minutes setup                     |
| `OP20_R`  | `10`            | 10 minutes run                       |
| `OP20_RN` | `WELD PER DWG`  | Routing note                         |
| `OP30`    | `N140`          | Work center = Outside Process (OP30) |
| `OP30_S`  | `0`             | No setup                             |
| `OP30_R`  | `0`             | No run                               |
| `OP30_RN` | `POWDER COAT`   | Routing note                         |

### Available Work Centers (WeldOps.txt)

The ComboBox values for assembly work centers are loaded from `WeldOps.txt`.
Typical values include: F400 (weld), N140 (outside process), etc.
