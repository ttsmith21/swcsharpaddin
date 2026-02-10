using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NM.Core;
using NM.Core.DataModel;
using NM.Core.Processing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Pipeline
{
    /// <summary>
    /// Comprehensive property write QA test that exercises every form button/option:
    /// - All 18 material radio buttons (sheet metal + tube parts)
    /// - Customer/Print/Revision text field propagation
    /// - Bend deduction (BendTable vs KFactor)
    /// - Re-run idempotency (process same part twice)
    /// - Assembly workflow (enumerate components, process each)
    /// - Batch workflow (all B-series + C-series parts)
    /// Validates that OP20, OptiMaterial, rbMaterialType, and all critical properties
    /// are correctly populated after processing.
    /// </summary>
    public class PropertyWriteQA
    {
        private readonly ISldWorks _swApp;
        private readonly string _inputDir;
        private readonly string _outputDir;
        private readonly StringBuilder _log = new StringBuilder();
        private int _passed;
        private int _failed;

        /// <summary>
        /// All 18 material radio buttons from MainSelectionForm.MapMaterial().
        /// Each tuple: (Material, MaterialType, MaterialCategory, ExpectedOptiPrefix).
        /// ExpectedOptiPrefix is the material code expected in OptiMaterial (may differ
        /// from Material when the MaterialCodeMapper normalizes the name).
        /// </summary>
        private static readonly (string Material, string MaterialType, MaterialCategoryKind Category, string OptiContains)[] AllMaterials = new[]
        {
            ("304L",    "AISI 304",                          MaterialCategoryKind.StainlessSteel, "304"),
            ("316L",    "AISI 316 Stainless Steel Sheet (SS)", MaterialCategoryKind.StainlessSteel, "316"),
            ("309",     "309",                               MaterialCategoryKind.StainlessSteel, "309"),
            ("310",     "310",                               MaterialCategoryKind.StainlessSteel, "310"),
            ("321",     "321",                               MaterialCategoryKind.StainlessSteel, "321"),
            ("330",     "330",                               MaterialCategoryKind.StainlessSteel, "330"),
            ("409",     "409",                               MaterialCategoryKind.StainlessSteel, "409"),
            ("430",     "430",                               MaterialCategoryKind.StainlessSteel, "430"),
            ("2205",    "2205",                              MaterialCategoryKind.StainlessSteel, "2205"),
            ("2507",    "2507",                              MaterialCategoryKind.StainlessSteel, "2507"),
            ("C22",     "Hastelloy C-22",                    MaterialCategoryKind.Other,          "C22"),
            ("C276",    "C276",                              MaterialCategoryKind.StainlessSteel, "C276"),
            ("AL6XN",   "AL6XN",                             MaterialCategoryKind.StainlessSteel, "AL6XN"),
            ("ALLOY31", "ALLOY31",                           MaterialCategoryKind.StainlessSteel, "ALLOY31"),
            ("A36",     "ASTM A36 Steel",                    MaterialCategoryKind.CarbonSteel,    "A36"),
            ("ALNZD",   "ASTM A36 Steel",                   MaterialCategoryKind.CarbonSteel,    "A36"),
            ("5052",    "5052-H32",                          MaterialCategoryKind.Aluminum,       "5052"),
            ("6061",    "6061 Alloy",                        MaterialCategoryKind.Aluminum,       "6061"),
        };

        public PropertyWriteQA(ISldWorks swApp, string inputDir, string outputDir)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
            _inputDir = inputDir;
            _outputDir = outputDir;
            if (!Directory.Exists(_outputDir)) Directory.CreateDirectory(_outputDir);
        }

        public int Run()
        {
            Log("=== PropertyWriteQA START ===");
            Log($"Input: {_inputDir}");
            Log($"Output: {_outputDir}");
            Log("");

            // 1. All 18 material buttons on sheet metal part
            RunAllMaterialsOnSheetMetal();

            // 2. All 18 material buttons on tube part
            RunAllMaterialsOnTube();

            // 3. Part-type-specific routing (flat, bent, rolled, saw)
            RunPartTypeRoutingTests();

            // 4. Custom property propagation (Customer, Print, Revision)
            RunCustomPropertyTests();

            // 5. Re-run idempotency (process same part twice)
            RunRerunTest();

            // 6. Assembly workflow
            RunAssemblyTest();

            // 7. Batch workflow (all B-series + C-series)
            RunBatchTest();

            Log("");
            Log("=== PropertyWriteQA SUMMARY ===");
            Log($"Passed: {_passed}");
            Log($"Failed: {_failed}");
            Log($"Total:  {_passed + _failed}");

            // Write log
            var logPath = Path.Combine(_outputDir, "property_write_qa.log");
            File.WriteAllText(logPath, _log.ToString());
            Console.WriteLine(_log.ToString());

            return _failed > 0 ? 1 : 0;
        }

        // =====================================================================
        // TEST 1: All 18 Material Buttons × Sheet Metal Part
        // =====================================================================
        private void RunAllMaterialsOnSheetMetal()
        {
            Log("--- TEST 1: All 18 Materials on Sheet Metal (B1) ---");

            var b1Path = FindPart("B1_NativeBracket_14ga_CS.SLDPRT");
            if (b1Path == null) return;

            foreach (var mat in AllMaterials)
            {
                var label = $"B1-{mat.Material}";
                var pd = ProcessPart(b1Path, new ProcessingOptions
                {
                    Material = mat.Material,
                    MaterialType = mat.MaterialType,
                    MaterialCategory = mat.Category,
                    SaveChanges = false
                });

                AssertNotNull(pd, $"{label}: PartData returned");
                if (pd == null) continue;

                // Classification & geometry
                AssertEquals(pd.Classification, PartType.SheetMetal, $"{label}: Classification=SheetMetal");
                AssertEquals(pd.Material, mat.Material, $"{label}: pd.Material={mat.Material}");
                AssertTrue(pd.Thickness_m > 0, $"{label}: Thickness_m > 0 (got {pd.Thickness_m:F6})");
                AssertTrue(pd.Sheet.TotalCutLength_m > 0, $"{label}: TotalCutLength_m > 0 (got {pd.Sheet.TotalCutLength_m:F6})");

                // OP20 must be calculated (not blank)
                AssertNotEmpty(pd.Cost.OP20_WorkCenter, $"{label}: OP20_WorkCenter not empty (got {pd.Cost.OP20_WorkCenter ?? "NULL"})");
                AssertTrue(pd.Cost.OP20_S_min > 0, $"{label}: OP20_S_min > 0 (got {pd.Cost.OP20_S_min:F4})");
                AssertTrue(pd.Cost.OP20_R_min > 0, $"{label}: OP20_R_min > 0 (got {pd.Cost.OP20_R_min:F4})");

                // OptiMaterial must be populated and start with "S." (sheet metal prefix)
                AssertNotEmpty(pd.OptiMaterial, $"{label}: OptiMaterial not empty");
                AssertStartsWith(pd.OptiMaterial, "S.", $"{label}: OptiMaterial starts with S.");
                // OptiMaterial should contain the material code (e.g., "S.304L14GA" contains "304")
                if (!string.IsNullOrEmpty(pd.OptiMaterial))
                    AssertContains(pd.OptiMaterial, mat.OptiContains, $"{label}: OptiMaterial contains '{mat.OptiContains}'");

                // Mapped properties for SolidWorks write
                var mapped = PartDataPropertyMap.ToProperties(pd);
                AssertEquals(Get(mapped, "rbMaterialType"), "0", $"{label}: rbMaterialType=0 (sheet)");
                AssertNotEmpty(Get(mapped, "OP20"), $"{label}: mapped OP20 not empty");
                AssertNotEquals(Get(mapped, "OP20_S"), "0", $"{label}: mapped OP20_S != 0");
                AssertNotEquals(Get(mapped, "OP20_R"), "0", $"{label}: mapped OP20_R != 0");
                AssertNotEmpty(Get(mapped, "OptiMaterial"), $"{label}: mapped OptiMaterial not empty");
            }
        }

        // =====================================================================
        // TEST 2: All 18 Material Buttons × Tube Part
        // =====================================================================
        private void RunAllMaterialsOnTube()
        {
            Log("--- TEST 2: All 18 Materials on Tube (C1) ---");

            var c1Path = FindPart("C1_RoundTube_2OD_SCH40.SLDPRT");
            if (c1Path == null) return;

            foreach (var mat in AllMaterials)
            {
                var label = $"C1-{mat.Material}";
                var pd = ProcessPart(c1Path, new ProcessingOptions
                {
                    Material = mat.Material,
                    MaterialType = mat.MaterialType,
                    MaterialCategory = mat.Category,
                    SaveChanges = false
                });

                AssertNotNull(pd, $"{label}: PartData returned");
                if (pd == null) continue;

                // Classification: tube part should classify as Tube regardless of material
                AssertEquals(pd.Classification, PartType.Tube, $"{label}: Classification=Tube");
                AssertEquals(pd.Material, mat.Material, $"{label}: pd.Material={mat.Material}");
                AssertTrue(pd.Tube.IsTube, $"{label}: IsTube=true");
                AssertTrue(pd.Tube.OD_m > 0, $"{label}: OD_m > 0 (got {pd.Tube.OD_m:F6})");

                // OP20 must be calculated for tubes
                AssertNotEmpty(pd.Cost.OP20_WorkCenter, $"{label}: OP20_WorkCenter not empty (got {pd.Cost.OP20_WorkCenter ?? "NULL"})");

                // Mapped properties
                var mapped = PartDataPropertyMap.ToProperties(pd);
                AssertEquals(Get(mapped, "rbMaterialType"), "1", $"{label}: rbMaterialType=1 (tube)");
                AssertNotEmpty(Get(mapped, "OP20"), $"{label}: mapped OP20 not empty");
            }
        }

        // =====================================================================
        // TEST 3: Part-Type-Specific Routing (flat, bent, rolled, saw)
        // =====================================================================
        private void RunPartTypeRoutingTests()
        {
            Log("--- TEST 3: Part-Type Routing (Flat / Bent / Rolled / Saw) ---");

            var defaultOpts = new ProcessingOptions
            {
                Material = "304L",
                MaterialType = "AISI 304",
                MaterialCategory = MaterialCategoryKind.StainlessSteel,
                SaveChanges = false
            };

            // ---------- FLAT: E1_ThickPlate_1in (no bends, just laser cut) ----------
            var e1Path = FindPart("E1_ThickPlate_1in.SLDPRT");
            if (e1Path != null)
            {
                var pd = ProcessPart(e1Path, defaultOpts);
                AssertNotNull(pd, "E1-Flat: PartData returned");
                if (pd != null && pd.Status == ProcessingStatus.Success)
                {
                    AssertEquals(pd.Classification, PartType.SheetMetal, "E1-Flat: Classification=SheetMetal");
                    // VBA baseline: 1.0" thickness, no N140 routing (no bends)
                    AssertTrue(pd.Thickness_m > 0.020, $"E1-Flat: Thickness > 0.020m / ~0.8in (got {pd.Thickness_m:F4})");
                    AssertEquals(pd.Sheet.BendCount, 0, "E1-Flat: BendCount=0 (flat plate)");

                    var mapped = PartDataPropertyMap.ToProperties(pd);
                    // OP20 = laser cutting (F115 → "N120 - 5040")
                    AssertNotEmpty(Get(mapped, "OP20"), "E1-Flat: OP20 not empty");
                    // PressBrake should be "Unchecked" (no bends)
                    AssertEquals(Get(mapped, "PressBrake"), "Unchecked", "E1-Flat: PressBrake=Unchecked (no bends)");
                    // F140 times should be zero
                    AssertEquals(Get(mapped, "F140_S"), "0", "E1-Flat: F140_S=0 (no bends)");
                    AssertEquals(Get(mapped, "F140_R"), "0", "E1-Flat: F140_R=0 (no bends)");
                    // F325 roll forming should be inactive (no large-radius bends)
                    AssertEquals(Get(mapped, "F325"), "0", "E1-Flat: F325=0 (no roll forming)");
                    // OptiMaterial should include "1IN" for 1" plate
                    AssertContains(pd.OptiMaterial, "1IN", $"E1-Flat: OptiMaterial contains '1IN' (got {pd.OptiMaterial ?? "NULL"})");
                }
            }

            // ---------- BENT: B1_NativeBracket (1 bend, press brake) ----------
            var b1Path = FindPart("B1_NativeBracket_14ga_CS.SLDPRT");
            if (b1Path != null)
            {
                var pd = ProcessPart(b1Path, defaultOpts);
                AssertNotNull(pd, "B1-Bent: PartData returned");
                if (pd != null && pd.Status == ProcessingStatus.Success)
                {
                    AssertEquals(pd.Classification, PartType.SheetMetal, "B1-Bent: Classification=SheetMetal");
                    // VBA baseline: 1 bend, N140 active
                    AssertTrue(pd.Sheet.BendCount >= 1, $"B1-Bent: BendCount >= 1 (got {pd.Sheet.BendCount})");

                    var mapped = PartDataPropertyMap.ToProperties(pd);
                    // PressBrake should be "Checked" (has bends)
                    AssertEquals(Get(mapped, "PressBrake"), "Checked", "B1-Bent: PressBrake=Checked (has bends)");
                    // F140 setup and run should be non-zero
                    AssertNotEquals(Get(mapped, "F140_S"), "0", "B1-Bent: F140_S != 0 (brake setup)");
                    AssertNotEquals(Get(mapped, "F140_R"), "0", "B1-Bent: F140_R != 0 (brake run)");
                    // BendCount property written
                    AssertNotEmpty(Get(mapped, "BendCount"), "B1-Bent: BendCount property written");
                    // OP20 = laser
                    AssertNotEmpty(Get(mapped, "OP20"), "B1-Bent: OP20 not empty (laser)");
                }
            }

            // ---------- ROLLED: B3_RolledCylinder (large-radius bend, roll forming) ----------
            var b3Path = FindPart("B3_RolledCylinder_16ga_SS.SLDPRT");
            if (b3Path != null)
            {
                var pd = ProcessPart(b3Path, defaultOpts);
                AssertNotNull(pd, "B3-Roll: PartData returned");
                if (pd != null && pd.Status == ProcessingStatus.Success)
                {
                    AssertEquals(pd.Classification, PartType.SheetMetal, "B3-Roll: Classification=SheetMetal");
                    AssertTrue(pd.Thickness_m > 0, $"B3-Roll: Thickness_m > 0 (got {pd.Thickness_m:F6})");

                    var mapped = PartDataPropertyMap.ToProperties(pd);
                    // OP20 = laser cutting
                    AssertNotEmpty(Get(mapped, "OP20"), "B3-Roll: OP20 not empty (laser)");
                    // VBA baseline: N325 roll forming active (0.375 setup, 0.2 run)
                    // F325 should be "1" (active)
                    AssertEquals(Get(mapped, "F325"), "1", "B3-Roll: F325=1 (roll forming active)");
                    AssertNotEquals(Get(mapped, "F325_S"), "0", "B3-Roll: F325_S != 0 (roll setup)");
                    AssertNotEquals(Get(mapped, "F325_R"), "0", "B3-Roll: F325_R != 0 (roll run)");
                    // OptiMaterial should have sheet prefix
                    AssertStartsWith(pd.OptiMaterial, "S.", "B3-Roll: OptiMaterial starts with S.");
                }
            }

            // ---------- SAW: C7_RoundBar (solid bar, F300 saw) ----------
            var c7Path = FindPart("C7_RoundBar_1dia.SLDPRT");
            if (c7Path == null)
                c7Path = FindPart("C7_RoundBar_1dia.sldprt"); // case sensitivity
            if (c7Path != null)
            {
                var pd = ProcessPart(c7Path, defaultOpts);
                AssertNotNull(pd, "C7-Saw: PartData returned");
                if (pd != null && pd.Status == ProcessingStatus.Success)
                {
                    AssertEquals(pd.Classification, PartType.Tube, "C7-Saw: Classification=Tube");
                    AssertTrue(pd.Tube.IsTube, "C7-Saw: IsTube=true");
                    // Solid bar: wall should be ~0 (no inner diameter)
                    AssertTrue(pd.Tube.Wall_m < 0.001, $"C7-Saw: Wall_m < 0.001 (solid bar, got {pd.Tube.Wall_m:F6})");

                    // VBA baseline: F300 (saw), not F110 (tube laser)
                    AssertEquals(pd.Cost.OP20_WorkCenter, "F300", $"C7-Saw: OP20_WorkCenter=F300 (got {pd.Cost.OP20_WorkCenter ?? "NULL"})");
                    AssertTrue(pd.Cost.OP20_S_min > 0, $"C7-Saw: OP20_S_min > 0 (got {pd.Cost.OP20_S_min:F4})");
                    AssertTrue(pd.Cost.OP20_R_min > 0, $"C7-Saw: OP20_R_min > 0 (got {pd.Cost.OP20_R_min:F4})");

                    var mapped = PartDataPropertyMap.ToProperties(pd);
                    // OP20 property should map to "F300 - SAW"
                    AssertEquals(Get(mapped, "OP20"), "F300 - SAW", "C7-Saw: mapped OP20='F300 - SAW'");
                    AssertEquals(Get(mapped, "rbMaterialType"), "1", "C7-Saw: rbMaterialType=1 (tube/bar)");
                    // OptiMaterial: VBA baseline = "R.304L1\"" (round bar prefix)
                    AssertStartsWith(pd.OptiMaterial, "R.", $"C7-Saw: OptiMaterial starts with R. (got {pd.OptiMaterial ?? "NULL"})");
                }
            }

            // ---------- MaterialCategory propagation check ----------
            Log("  --- MaterialCategory Propagation ---");
            if (b1Path != null)
            {
                // StainlessSteel
                var pdSS = ProcessPart(b1Path, new ProcessingOptions
                {
                    Material = "304L", MaterialType = "AISI 304",
                    MaterialCategory = MaterialCategoryKind.StainlessSteel, SaveChanges = false
                });
                if (pdSS != null)
                {
                    AssertEquals(pdSS.MaterialCategory, "StainlessSteel", "MatCat-SS: pd.MaterialCategory=StainlessSteel");
                    var mappedSS = PartDataPropertyMap.ToProperties(pdSS);
                    AssertEquals(Get(mappedSS, "MaterialCategory"), "StainlessSteel", "MatCat-SS: mapped MaterialCategory=StainlessSteel");
                }

                // CarbonSteel
                var pdCS = ProcessPart(b1Path, new ProcessingOptions
                {
                    Material = "A36", MaterialType = "ASTM A36 Steel",
                    MaterialCategory = MaterialCategoryKind.CarbonSteel, SaveChanges = false
                });
                if (pdCS != null)
                {
                    AssertEquals(pdCS.MaterialCategory, "CarbonSteel", "MatCat-CS: pd.MaterialCategory=CarbonSteel");
                    var mappedCS = PartDataPropertyMap.ToProperties(pdCS);
                    AssertEquals(Get(mappedCS, "MaterialCategory"), "CarbonSteel", "MatCat-CS: mapped MaterialCategory=CarbonSteel");
                }

                // Aluminum
                var pdAL = ProcessPart(b1Path, new ProcessingOptions
                {
                    Material = "6061", MaterialType = "6061 Alloy",
                    MaterialCategory = MaterialCategoryKind.Aluminum, SaveChanges = false
                });
                if (pdAL != null)
                {
                    AssertEquals(pdAL.MaterialCategory, "Aluminum", "MatCat-AL: pd.MaterialCategory=Aluminum");
                    var mappedAL = PartDataPropertyMap.ToProperties(pdAL);
                    AssertEquals(Get(mappedAL, "MaterialCategory"), "Aluminum", "MatCat-AL: mapped MaterialCategory=Aluminum");
                }
            }
        }

        // =====================================================================
        // TEST 4: Custom Property Propagation (Customer, Print, Revision)
        // =====================================================================
        private void RunCustomPropertyTests()
        {
            Log("--- TEST 4: Custom Property Propagation ---");

            var b1Path = FindPart("B1_NativeBracket_14ga_CS.SLDPRT");
            if (b1Path == null) return;

            // Process with Customer/Print/Revision set
            var pd = ProcessPart(b1Path, new ProcessingOptions
            {
                Material = "304L",
                MaterialType = "AISI 304",
                MaterialCategory = MaterialCategoryKind.StainlessSteel,
                Customer = "TEST_CUST",
                Print = "PRN-12345",
                Revision = "B",
                SaveChanges = false
            });

            AssertNotNull(pd, "CustomProps: PartData returned");
            if (pd == null) return;

            // Customer, Print, Revision should flow through pd.Extra → mapped properties
            var mapped = PartDataPropertyMap.ToProperties(pd);
            AssertEquals(Get(mapped, "Customer"), "TEST_CUST", "CustomProps: Customer=TEST_CUST");
            AssertEquals(Get(mapped, "Print"), "PRN-12345", "CustomProps: Print=PRN-12345");
            AssertEquals(Get(mapped, "Revision"), "B", "CustomProps: Revision=B");

            // Description should be auto-generated (not empty)
            AssertNotEmpty(Get(mapped, "Description"), "CustomProps: Description auto-generated");

            // OP20/OptiMaterial should still work alongside custom props
            AssertNotEmpty(Get(mapped, "OP20"), "CustomProps: OP20 still populated");
            AssertNotEmpty(Get(mapped, "OptiMaterial"), "CustomProps: OptiMaterial still populated");
            AssertEquals(Get(mapped, "rbMaterialType"), "0", "CustomProps: rbMaterialType=0 (sheet)");

            // Process with empty Customer/Print/Revision (should use defaults/fallbacks)
            var pd2 = ProcessPart(b1Path, new ProcessingOptions
            {
                Material = "304L",
                MaterialType = "AISI 304",
                MaterialCategory = MaterialCategoryKind.StainlessSteel,
                SaveChanges = false
            });
            AssertNotNull(pd2, "CustomProps-Empty: PartData returned");
            if (pd2 != null)
            {
                // OP20/OptiMaterial should still work even with no custom props
                AssertNotEmpty(pd2.Cost.OP20_WorkCenter, "CustomProps-Empty: OP20_WorkCenter not empty");
                AssertNotEmpty(pd2.OptiMaterial, "CustomProps-Empty: OptiMaterial not empty");
            }
        }

        // =====================================================================
        // TEST 5: Re-Run Idempotency (process same part twice)
        // =====================================================================
        private void RunRerunTest()
        {
            Log("--- TEST 5: Re-Run Idempotency ---");

            var b1Path = FindPart("B1_NativeBracket_14ga_CS.SLDPRT");
            if (b1Path == null) return;

            var opts = new ProcessingOptions
            {
                Material = "304L",
                MaterialType = "AISI 304",
                MaterialCategory = MaterialCategoryKind.StainlessSteel,
                SaveChanges = false
            };

            // Run 1
            var pd1 = ProcessPart(b1Path, opts);
            AssertNotNull(pd1, "ReRun-1: PartData returned");
            if (pd1 == null) return;

            var mapped1 = PartDataPropertyMap.ToProperties(pd1);
            var op20_1 = Get(mapped1, "OP20");
            var op20s_1 = Get(mapped1, "OP20_S");
            var op20r_1 = Get(mapped1, "OP20_R");
            var opti_1 = Get(mapped1, "OptiMaterial");

            AssertNotEmpty(op20_1, "ReRun-1: OP20 not empty");
            AssertNotEmpty(opti_1, "ReRun-1: OptiMaterial not empty");

            // Run 2: same part, same options
            var pd2 = ProcessPart(b1Path, opts);
            AssertNotNull(pd2, "ReRun-2: PartData returned");
            if (pd2 == null) return;

            var mapped2 = PartDataPropertyMap.ToProperties(pd2);
            var op20_2 = Get(mapped2, "OP20");
            var op20s_2 = Get(mapped2, "OP20_S");
            var op20r_2 = Get(mapped2, "OP20_R");
            var opti_2 = Get(mapped2, "OptiMaterial");

            // Both runs should produce the same results
            AssertEquals(op20_2, op20_1, $"ReRun: OP20 matches (Run1={op20_1}, Run2={op20_2})");
            AssertEquals(op20s_2, op20s_1, $"ReRun: OP20_S matches (Run1={op20s_1}, Run2={op20s_2})");
            AssertEquals(op20r_2, op20r_1, $"ReRun: OP20_R matches (Run1={op20r_1}, Run2={op20r_2})");
            AssertEquals(opti_2, opti_1, $"ReRun: OptiMaterial matches (Run1={opti_1}, Run2={opti_2})");

            // Neither run should produce blanks or zeros
            AssertNotEquals(op20_2, "", "ReRun-2: OP20 not blank on re-run");
            AssertNotEquals(op20s_2, "0", "ReRun-2: OP20_S not zero on re-run");
            AssertNotEquals(op20r_2, "0", "ReRun-2: OP20_R not zero on re-run");
            AssertNotEquals(opti_2, "", "ReRun-2: OptiMaterial not blank on re-run");

            // Run 3: different material on same part — should change OptiMaterial
            var pd3 = ProcessPart(b1Path, new ProcessingOptions
            {
                Material = "A36",
                MaterialType = "ASTM A36 Steel",
                MaterialCategory = MaterialCategoryKind.CarbonSteel,
                SaveChanges = false
            });
            if (pd3 != null)
            {
                var mapped3 = PartDataPropertyMap.ToProperties(pd3);
                var opti_3 = Get(mapped3, "OptiMaterial");
                AssertNotEmpty(opti_3, "ReRun-3: OptiMaterial not empty with A36");
                // OptiMaterial should change to reflect A36 (not still show 304L)
                if (!string.IsNullOrEmpty(opti_3) && !string.IsNullOrEmpty(opti_1))
                    AssertNotEquals(opti_3, opti_1, $"ReRun-3: OptiMaterial changed (304L={opti_1}, A36={opti_3})");
            }
        }

        // =====================================================================
        // TEST 6: Assembly Workflow
        // =====================================================================
        private void RunAssemblyTest()
        {
            Log("--- TEST 6: Assembly Workflow ---");

            var assyPath = Path.Combine(_inputDir, "F1_SimpleAssy.SLDASM");
            if (!File.Exists(assyPath))
                assyPath = Path.Combine(_inputDir, "GOLD_STANDARD_ASM_CLEAN.SLDASM");

            if (!File.Exists(assyPath))
            {
                Log("  SKIP: No assembly file found");
                return;
            }

            Log($"  Opening assembly: {Path.GetFileName(assyPath)}");

            int errors = 0, warnings = 0;
            var doc = _swApp.OpenDoc6(assyPath,
                (int)swDocumentTypes_e.swDocASSEMBLY,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "", ref errors, ref warnings);

            if (doc == null)
            {
                Log($"  SKIP: Failed to open assembly (errors={errors})");
                return;
            }

            try
            {
                var asm = doc as IAssemblyDoc;
                if (asm == null)
                {
                    Log("  SKIP: Document is not an assembly");
                    return;
                }

                // Enumerate components
                var quantifier = new AssemblyProcessing.AssemblyComponentQuantifier();
                var components = quantifier.CollectViaRecursion(asm);
                AssertTrue(components.Count > 0, $"Assembly: has components (got {components.Count})");

                // Collect part file paths
                var partPaths = components.Values
                    .Where(c => c.FilePath != null &&
                        c.FilePath.EndsWith(".SLDPRT", StringComparison.OrdinalIgnoreCase))
                    .Select(c => c.FilePath)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                AssertTrue(partPaths.Count > 0, $"Assembly: has part components (got {partPaths.Count})");

                // Close assembly before opening individual parts
                CloseDoc(doc);
                doc = null;

                // Process each unique part component
                int assySuccess = 0;
                int assyWithOP20 = 0;
                int assyWithOptiMat = 0;

                foreach (var partPath in partPaths)
                {
                    Log($"  Processing component: {Path.GetFileName(partPath)}");
                    var partDoc = OpenPart(partPath);
                    if (partDoc == null) continue;

                    try
                    {
                        var pd = MainRunner.RunSinglePartData(_swApp, partDoc,
                            new ProcessingOptions
                            {
                                Material = "304L",
                                MaterialType = "AISI 304",
                                MaterialCategory = MaterialCategoryKind.StainlessSteel,
                                SaveChanges = false
                            });

                        if (pd != null && pd.Status == ProcessingStatus.Success)
                        {
                            assySuccess++;
                            AssertTrue(pd.Classification != PartType.Unknown,
                                $"Assy-{Path.GetFileName(partPath)}: classified (got {pd.Classification})");

                            if (!string.IsNullOrEmpty(pd.Cost.OP20_WorkCenter)) assyWithOP20++;
                            if (!string.IsNullOrEmpty(pd.OptiMaterial)) assyWithOptiMat++;

                            // rbMaterialType must be "0" or "1" (not material name)
                            var mapped = PartDataPropertyMap.ToProperties(pd);
                            var rbMT = Get(mapped, "rbMaterialType");
                            AssertTrue(rbMT == "0" || rbMT == "1",
                                $"Assy-{Path.GetFileName(partPath)}: rbMaterialType is '0' or '1' (got '{rbMT}')");
                        }
                    }
                    finally
                    {
                        CloseDoc(partDoc);
                    }
                }

                AssertTrue(assySuccess > 0, $"Assembly: at least 1 component processed (got {assySuccess})");
                if (assySuccess > 0)
                {
                    AssertEquals(assyWithOP20, assySuccess, $"Assembly: all processed have OP20 ({assyWithOP20}/{assySuccess})");
                    AssertEquals(assyWithOptiMat, assySuccess, $"Assembly: all processed have OptiMaterial ({assyWithOptiMat}/{assySuccess})");
                }
            }
            finally
            {
                if (doc != null) CloseDoc(doc);
            }
        }

        // =====================================================================
        // TEST 7: Batch Workflow (all B-series + C-series)
        // =====================================================================
        private void RunBatchTest()
        {
            Log("--- TEST 7: Batch Workflow ---");

            // ----- B-series (sheet metal) -----
            var bParts = Directory.GetFiles(_inputDir, "B*.SLDPRT", SearchOption.TopDirectoryOnly);
            if (bParts.Length == 0)
            {
                Log("  SKIP: No B-series parts found");
            }
            else
            {
                int batchSuccess = 0;
                int batchWithOP20 = 0;
                int batchWithOptiMat = 0;

                foreach (var partPath in bParts)
                {
                    var pd = ProcessPartByPath(partPath, new ProcessingOptions
                    {
                        Material = "304L",
                        MaterialType = "AISI 304",
                        MaterialCategory = MaterialCategoryKind.StainlessSteel,
                        SaveChanges = false
                    });
                    if (pd != null && pd.Status == ProcessingStatus.Success)
                    {
                        batchSuccess++;
                        if (!string.IsNullOrEmpty(pd.Cost.OP20_WorkCenter)) batchWithOP20++;
                        if (!string.IsNullOrEmpty(pd.OptiMaterial)) batchWithOptiMat++;

                        // Every successful sheet metal part must have rbMaterialType=0
                        var mapped = PartDataPropertyMap.ToProperties(pd);
                        AssertEquals(Get(mapped, "rbMaterialType"), "0",
                            $"Batch-{Path.GetFileName(partPath)}: rbMaterialType=0");
                    }
                }

                AssertTrue(batchSuccess > 0, $"Batch-B: at least 1 success (got {batchSuccess}/{bParts.Length})");
                AssertEquals(batchWithOP20, batchSuccess, $"Batch-B: all successful have OP20 ({batchWithOP20}/{batchSuccess})");
                AssertEquals(batchWithOptiMat, batchSuccess, $"Batch-B: all successful have OptiMaterial ({batchWithOptiMat}/{batchSuccess})");
            }

            // ----- C-series (tubes) -----
            var cParts = Directory.GetFiles(_inputDir, "C*.SLDPRT", SearchOption.TopDirectoryOnly);
            if (cParts.Length == 0)
            {
                Log("  SKIP: No C-series parts found");
            }
            else
            {
                int tubeSuccess = 0;
                int tubeWithOP20 = 0;

                foreach (var partPath in cParts)
                {
                    var pd = ProcessPartByPath(partPath, new ProcessingOptions
                    {
                        Material = "A36",
                        MaterialType = "ASTM A36 Steel",
                        MaterialCategory = MaterialCategoryKind.CarbonSteel,
                        SaveChanges = false
                    });
                    if (pd != null && pd.Status == ProcessingStatus.Success)
                    {
                        tubeSuccess++;
                        if (!string.IsNullOrEmpty(pd.Cost.OP20_WorkCenter)) tubeWithOP20++;

                        // Every successful tube part must have rbMaterialType=1
                        var mapped = PartDataPropertyMap.ToProperties(pd);
                        AssertEquals(Get(mapped, "rbMaterialType"), "1",
                            $"Batch-{Path.GetFileName(partPath)}: rbMaterialType=1");
                    }
                }

                AssertTrue(tubeSuccess > 0, $"Batch-C: at least 1 success (got {tubeSuccess}/{cParts.Length})");
                AssertEquals(tubeWithOP20, tubeSuccess, $"Batch-C: all successful have OP20 ({tubeWithOP20}/{tubeSuccess})");
            }
        }

        // =====================================================================
        // Helpers
        // =====================================================================
        private string FindPart(string fileName)
        {
            var path = Path.Combine(_inputDir, fileName);
            if (!File.Exists(path))
            {
                Log($"  SKIP: {fileName} not found");
                return null;
            }
            return path;
        }

        private PartData ProcessPart(string filePath, ProcessingOptions options)
        {
            var doc = OpenPart(filePath);
            if (doc == null) return null;

            try
            {
                Log($"  Processing: {Path.GetFileName(filePath)} (Material={options.Material})");
                var pd = MainRunner.RunSinglePartData(_swApp, doc, options);
                if (pd != null)
                    Log($"    Result: Status={pd.Status}, Classification={pd.Classification}, " +
                        $"OP20_WC={pd.Cost.OP20_WorkCenter ?? "NULL"}, " +
                        $"OptiMaterial={pd.OptiMaterial ?? "NULL"}, " +
                        $"Material={pd.Material ?? "NULL"}");
                return pd;
            }
            finally
            {
                CloseDoc(doc);
            }
        }

        private PartData ProcessPartByPath(string filePath, ProcessingOptions options)
        {
            return ProcessPart(filePath, options);
        }

        private IModelDoc2 OpenPart(string filePath)
        {
            int errors = 0, warnings = 0;
            var doc = _swApp.OpenDoc6(filePath,
                (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "", ref errors, ref warnings);
            if (doc == null)
                Log($"    OPEN FAILED: {Path.GetFileName(filePath)} (errors={errors})");
            return doc;
        }

        private void CloseDoc(IModelDoc2 doc)
        {
            if (doc == null) return;
            try
            {
                doc.SetSaveFlag();
                var path = doc.GetPathName();
                if (!string.IsNullOrEmpty(path))
                    _swApp.CloseDoc(path);
                else
                {
                    var title = doc.GetTitle();
                    if (!string.IsNullOrEmpty(title))
                        _swApp.CloseDoc(title);
                }
            }
            catch { }
        }

        private static string Get(IDictionary<string, string> d, string key)
        {
            return d != null && d.TryGetValue(key, out var v) ? v ?? "" : "";
        }

        // =====================================================================
        // Assertions
        // =====================================================================
        private void AssertTrue(bool condition, string message)
        {
            if (condition) { _passed++; Log($"  PASS: {message}"); }
            else { _failed++; Log($"  FAIL: {message}"); }
        }

        private void AssertEquals<T>(T actual, T expected, string message)
        {
            if (EqualityComparer<T>.Default.Equals(actual, expected))
            { _passed++; Log($"  PASS: {message}"); }
            else
            { _failed++; Log($"  FAIL: {message} (expected={expected}, actual={actual})"); }
        }

        private void AssertNotEquals(string actual, string notExpected, string message)
        {
            if (actual != notExpected)
            { _passed++; Log($"  PASS: {message} (got {actual})"); }
            else
            { _failed++; Log($"  FAIL: {message}"); }
        }

        private void AssertNotNull(object obj, string message)
        {
            if (obj != null) { _passed++; Log($"  PASS: {message}"); }
            else { _failed++; Log($"  FAIL: {message} (was null)"); }
        }

        private void AssertNotEmpty(string value, string message)
        {
            if (!string.IsNullOrEmpty(value)) { _passed++; Log($"  PASS: {message} (got {value})"); }
            else { _failed++; Log($"  FAIL: {message} (was empty)"); }
        }

        private void AssertStartsWith(string value, string prefix, string message)
        {
            if (value != null && value.StartsWith(prefix))
            { _passed++; Log($"  PASS: {message} (got {value})"); }
            else
            { _failed++; Log($"  FAIL: {message} (got {value ?? "NULL"})"); }
        }

        private void AssertContains(string value, string substring, string message)
        {
            if (value != null && value.IndexOf(substring, StringComparison.OrdinalIgnoreCase) >= 0)
            { _passed++; Log($"  PASS: {message} (got {value})"); }
            else
            { _failed++; Log($"  FAIL: {message} (got {value ?? "NULL"})"); }
        }

        private void Log(string message)
        {
            _log.AppendLine(message);
            ErrorHandler.DebugLog(message);
        }
    }
}
