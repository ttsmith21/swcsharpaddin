# Quick Status - What Works Right Now

## ? WORKS (Test These!)

### Sheet Metal Conversion
```
1. Open SolidWorks
2. Open a simple flat solid part (like a machined block)
3. NM Classifier ? "Run Single-Part Pipeline"
4. Select 304L + Bend Table
5. ? Part converts to sheet metal with bends
```

**Expected Output:**
- Sheet metal feature added
- Part validates successfully
- Logs show `[SMDBG]` diagnostics
- Save prompt (optional)

### Material Selection UI
```
All materials load correctly:
? 304L, 316L (Stainless Steel)
? A36 (Carbon Steel)
? 6061, 5052 (Aluminum)
? 409, 2205, 309, C22, AL6XN (Specialty)
? ALNZD, Other
```

### Validation
```
? Single solid body detection
? Multi-body rejection
? Empty geometry rejection
? Surface body rejection
```

### Infrastructure
```
? Logging to ErrorLog.txt
? Performance timing (if enabled in config)
? Call stack tracking
? Error handling with graceful degradation
```

---

## ?? DOESN'T WORK (Skip These for Now)

### Tube Parts
```
? Opening a tube/pipe part
? SimpleTubeProcessor.CanProcess() ? returns false
? No tube geometry extraction
? ExternalStart won't trigger
```

**Current Behavior:**
- Tube parts are treated as "fallback" (no processing)
- `[TUBE]` logs show "STUBBED - returning false"
- All tube services exist but aren't called

**Fix Location:**
`src\NM.Core\Processing\SimpleTubeProcessor.cs` ? `ExtractTubeGeometry()`

### Custom Properties
```
? No custom properties written to parts
? Can't read IsSheetMetal, Thickness, OptiMaterial, etc.
? Manufacturing data (OP20, F140, etc.) not populated
```

**Current Behavior:**
- Properties button in UI is placeholder
- Processing works but doesn't save manufacturing data

**Fix Location:**
Epic 0 in `MasterPlan.md` - needs new service

### Manufacturing Intelligence
```
? No weight calculations
? No bend analysis (count, tonnage)
? No cost calculations
? No OptiMaterial lookup
```

**Current Behavior:**
- Phase 2 features not started

**Fix Location:**
Epics 6-9 in `MasterPlan.md`

---

## ?? Test Scenarios

### ? Scenario 1: Simple Sheet Metal (WORKS)
```
Part: Flat rectangular solid, 12" × 6" × 0.125" thick
Material: 304L
Bend Option: Bend Table (SS)
Expected: ? Converts successfully, adds sheet metal feature
```

### ? Scenario 2: Multi-body Rejection (WORKS)
```
Part: Two separate solid bodies
Expected: ? Validation fails with "Multi-body detected"
```

### ? Scenario 3: Empty Geometry (WORKS)
```
Part: Empty part document
Expected: ? Validation fails with "No solid bodies"
```

### ?? Scenario 4: Tube Part (BLOCKED)
```
Part: 2" OD × 0.125" wall × 12" long tube
Material: 304L
Expected: ?? Falls back to generic processor (no tube-specific logic)
Actual: Shows success but doesn't extract tube data
```

### ?? Scenario 5: Custom Properties (NOT IMPLEMENTED)
```
Part: Any processed part
Check: Custom Properties ? Configuration Specific
Expected: ?? No manufacturing properties written yet
```

---

## ?? Testing Checklist

### Before Handoff Testing
- [x] Sheet metal conversion works (304L + Bend Table)
- [x] Validation rejects multi-body parts
- [x] Validation rejects empty parts
- [x] UI shows all materials
- [x] Bend table path resolution works (O: drive + UNC)
- [x] Logging outputs to ErrorLog.txt
- [x] Performance tracking (if enabled)
- [ ] A36 + Carbon Steel bend table *(not tested yet)*
- [ ] K-Factor mode (6061/5052) *(not tested yet)*
- [ ] Already-sheet-metal graceful handling *(not tested yet)*

### Post-Handoff Testing (For Programmer)
- [ ] Tube geometry extraction (cylinder faces)
- [ ] Tube OD/ID/length calculations
- [ ] ExternalStart trigger (16"×0.5")
- [ ] Custom property writes (manufacturing data)
- [ ] Weight calculations
- [ ] Bend analysis
- [ ] Costing formulas

---

## ?? Common Gotchas

### 1. Build Warnings (IGNORE THESE)
```
?? MSB3026: Could not copy .dll (locked by SolidWorks)
? Normal during development; close SW before building
```

### 2. Bend Table Paths
```
? O:\Engineering Department\Solidworks\Sheet Metal Bend Tables\...
? Falls back to SW install folder if not found
? Tries both .xls and .xlsx
```

### 3. Unit Conversions
```
?? SolidWorks internal units = METERS
? Convert to inches for display/calculations
? See GeometryExtension.cs for helpers
```

### 4. Threading
```
?? All SolidWorks COM calls MUST be on main thread (STA)
? Don't use Task.Run() or background threads for SW API
```

### 5. COM Lifetime
```
?? Don't call Marshal.ReleaseComObject() everywhere
? Only use in high-volume loops
? Never release ISldWorks singleton
```

---

## ?? Questions & Support

### Where to Look
1. **Error logs:** Check `ErrorLog.txt` (configured path)
2. **Debug output:** Visual Studio Output window
3. **SolidWorks:** Check for error messages in status bar
4. **Master plan:** `MasterPlan.md` for roadmap
5. **Tube guide:** `docs\TUBE_HANDOFF.md` for debugging

### What to Check First
1. Is SolidWorks running? (locks DLL during build)
2. Is add-in loaded? (Tools ? Add-Ins ? NM SwAddin)
3. Are logs enabled? (Configuration.Logging.EnableDebugMode)
4. Is part type correct? (Must be Part, not Assembly)
5. Is part valid? (Single solid body, non-empty)

---

*Quick reference created: January 2025*
*For real programmer handoff - test sheet metal first!*
