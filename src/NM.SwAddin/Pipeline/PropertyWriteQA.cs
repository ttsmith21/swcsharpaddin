using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NM.Core;
using NM.Core.DataModel;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Pipeline
{
    /// <summary>
    /// Comprehensive property write QA test that exercises every workflow path:
    /// - Single sheet metal part with form options
    /// - Single tube part with form options
    /// - Assembly → per-component processing
    /// - Batch folder processing
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

            // ===== SINGLE PART: Sheet Metal =====
            RunSheetMetalTests();

            // ===== SINGLE PART: Tube =====
            RunTubeTests();

            // ===== ASSEMBLY WORKFLOW =====
            RunAssemblyTest();

            // ===== BATCH WORKFLOW =====
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
        // Sheet Metal Tests
        // =====================================================================
        private void RunSheetMetalTests()
        {
            Log("--- Sheet Metal Single-Part Tests ---");

            // B1: Native bracket, 14ga CS — the bread-and-butter case
            var b1 = FindPart("B1_NativeBracket_14ga_CS.SLDPRT");
            if (b1 != null)
            {
                // Test 1: Default options (304L)
                var pd1 = ProcessPart(b1, new ProcessingOptions
                {
                    Material = "304L",
                    MaterialType = "AISI 304",
                    SaveChanges = false
                });
                AssertNotNull(pd1, "B1-304L: PartData returned");
                if (pd1 != null)
                {
                    AssertEquals(pd1.Classification, PartType.SheetMetal, "B1-304L: Classification");
                    AssertTrue(pd1.Thickness_m > 0, $"B1-304L: Thickness_m > 0 (got {pd1.Thickness_m:F6})");
                    AssertTrue(pd1.Sheet.TotalCutLength_m > 0, $"B1-304L: TotalCutLength_m > 0 (got {pd1.Sheet.TotalCutLength_m:F6})");
                    AssertEquals(pd1.Cost.OP20_WorkCenter, "F115", "B1-304L: OP20_WorkCenter=F115");
                    AssertTrue(pd1.Cost.OP20_S_min > 0, $"B1-304L: OP20_S_min > 0 (got {pd1.Cost.OP20_S_min:F4})");
                    AssertTrue(pd1.Cost.OP20_R_min > 0, $"B1-304L: OP20_R_min > 0 (got {pd1.Cost.OP20_R_min:F4})");
                    AssertNotEmpty(pd1.OptiMaterial, "B1-304L: OptiMaterial not empty");
                    AssertStartsWith(pd1.OptiMaterial, "S.", "B1-304L: OptiMaterial starts with S.");

                    // Verify mapped properties
                    var mapped = PartDataPropertyMap.ToProperties(pd1);
                    AssertEquals(Get(mapped, "rbMaterialType"), "0", "B1-304L: rbMaterialType=0 (sheet)");
                    AssertNotEmpty(Get(mapped, "OP20"), "B1-304L: mapped OP20 not empty");
                    AssertNotEquals(Get(mapped, "OP20_S"), "0", "B1-304L: mapped OP20_S != 0");
                    AssertNotEquals(Get(mapped, "OP20_R"), "0", "B1-304L: mapped OP20_R != 0");
                    AssertNotEmpty(Get(mapped, "OptiMaterial"), "B1-304L: mapped OptiMaterial not empty");
                }

                // Test 2: Different material (316L)
                var pd2 = ProcessPart(b1, new ProcessingOptions
                {
                    Material = "316L",
                    MaterialType = "AISI 316",
                    SaveChanges = false
                });
                if (pd2 != null)
                {
                    AssertEquals(pd2.Cost.OP20_WorkCenter, "F115", "B1-316L: OP20_WorkCenter=F115");
                    AssertNotEmpty(pd2.OptiMaterial, "B1-316L: OptiMaterial not empty");
                    if (!string.IsNullOrEmpty(pd2.OptiMaterial))
                        AssertTrue(pd2.OptiMaterial.Contains("316"), $"B1-316L: OptiMaterial contains 316 (got {pd2.OptiMaterial})");
                }
            }

            // B2: Imported bracket (STEP-based)
            var b2 = FindPart("B2_ImportedBracket_14ga_CS.SLDPRT");
            if (b2 != null)
            {
                var pd = ProcessPart(b2, new ProcessingOptions { Material = "304L", MaterialType = "AISI 304", SaveChanges = false });
                if (pd != null)
                {
                    AssertEquals(pd.Classification, PartType.SheetMetal, "B2: Classification=SheetMetal");
                    AssertTrue(pd.Thickness_m > 0, $"B2: Thickness_m > 0 (got {pd.Thickness_m:F6})");
                    AssertEquals(pd.Cost.OP20_WorkCenter, "F115", "B2: OP20_WorkCenter=F115");
                }
            }
        }

        // =====================================================================
        // Tube Tests
        // =====================================================================
        private void RunTubeTests()
        {
            Log("--- Tube Single-Part Tests ---");

            // C1: Round tube
            var c1 = FindPart("C1_RoundTube_2OD_SCH40.SLDPRT");
            if (c1 != null)
            {
                var pd = ProcessPart(c1, new ProcessingOptions { Material = "A36 Steel", SaveChanges = false });
                if (pd != null)
                {
                    AssertEquals(pd.Classification, PartType.Tube, "C1: Classification=Tube");
                    AssertTrue(pd.Tube.IsTube, "C1: IsTube=true");
                    AssertTrue(pd.Tube.OD_m > 0, $"C1: OD_m > 0 (got {pd.Tube.OD_m:F6})");
                    AssertNotEmpty(pd.Cost.OP20_WorkCenter, $"C1: OP20_WorkCenter not empty (got {pd.Cost.OP20_WorkCenter ?? "NULL"})");

                    var mapped = PartDataPropertyMap.ToProperties(pd);
                    AssertEquals(Get(mapped, "rbMaterialType"), "1", "C1: rbMaterialType=1 (tube)");
                    AssertNotEmpty(Get(mapped, "OP20"), "C1: mapped OP20 not empty");
                }
            }

            // C2: Rectangular tube
            var c2 = FindPart("C2_RectTube_2x1.SLDPRT");
            if (c2 != null)
            {
                var pd = ProcessPart(c2, new ProcessingOptions { Material = "A36 Steel", SaveChanges = false });
                if (pd != null)
                {
                    AssertEquals(pd.Classification, PartType.Tube, "C2: Classification=Tube");
                    var mapped = PartDataPropertyMap.ToProperties(pd);
                    AssertEquals(Get(mapped, "rbMaterialType"), "1", "C2: rbMaterialType=1 (tube)");
                }
            }
        }

        // =====================================================================
        // Assembly Test
        // =====================================================================
        private void RunAssemblyTest()
        {
            Log("--- Assembly Workflow Test ---");

            // F1_SimpleAssy or GOLD_STANDARD_ASM_CLEAN
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

                // Process first part component
                var firstPart = components.Values
                    .FirstOrDefault(c => c.FilePath != null &&
                        c.FilePath.EndsWith(".SLDPRT", StringComparison.OrdinalIgnoreCase));

                if (firstPart != null)
                {
                    Log($"  Processing component: {Path.GetFileName(firstPart.FilePath)}");

                    // Close assembly, open part, process
                    CloseDoc(doc);
                    doc = null;

                    var partDoc = OpenPart(firstPart.FilePath);
                    if (partDoc != null)
                    {
                        try
                        {
                            var pd = MainRunner.RunSinglePartData(_swApp, partDoc,
                                new ProcessingOptions { Material = "304L", MaterialType = "AISI 304", SaveChanges = false });
                            AssertNotNull(pd, "Assembly-Part: PartData returned");
                            if (pd != null && pd.Status == ProcessingStatus.Success)
                            {
                                AssertTrue(pd.Classification != PartType.Unknown, $"Assembly-Part: classified (got {pd.Classification})");
                            }
                        }
                        finally
                        {
                            CloseDoc(partDoc);
                        }
                    }
                }
                else
                {
                    Log("  SKIP: No part components found in assembly");
                    CloseDoc(doc);
                    doc = null;
                }
            }
            finally
            {
                if (doc != null) CloseDoc(doc);
            }
        }

        // =====================================================================
        // Batch Test
        // =====================================================================
        private void RunBatchTest()
        {
            Log("--- Batch Workflow Test ---");

            // Process ALL B-series parts (sheet metal) as a batch
            var bParts = Directory.GetFiles(_inputDir, "B*.SLDPRT", SearchOption.TopDirectoryOnly);
            if (bParts.Length == 0)
            {
                Log("  SKIP: No B-series parts found");
                return;
            }

            int batchSuccess = 0;
            int batchWithOP20 = 0;
            int batchWithOptiMat = 0;

            foreach (var partPath in bParts)
            {
                var pd = ProcessPartByPath(partPath, new ProcessingOptions
                {
                    Material = "304L", MaterialType = "AISI 304", SaveChanges = false
                });
                if (pd != null && pd.Status == ProcessingStatus.Success)
                {
                    batchSuccess++;
                    if (!string.IsNullOrEmpty(pd.Cost.OP20_WorkCenter)) batchWithOP20++;
                    if (!string.IsNullOrEmpty(pd.OptiMaterial)) batchWithOptiMat++;
                }
            }

            AssertTrue(batchSuccess > 0, $"Batch: at least 1 B-series success (got {batchSuccess}/{bParts.Length})");
            AssertEquals(batchWithOP20, batchSuccess, $"Batch: all successful parts have OP20 ({batchWithOP20}/{batchSuccess})");
            AssertEquals(batchWithOptiMat, batchSuccess, $"Batch: all successful parts have OptiMaterial ({batchWithOptiMat}/{batchSuccess})");

            // Process ALL C-series parts (tubes) as a batch
            var cParts = Directory.GetFiles(_inputDir, "C*.SLDPRT", SearchOption.TopDirectoryOnly);
            if (cParts.Length > 0)
            {
                int tubeSuccess = 0;
                int tubeWithOP20 = 0;
                foreach (var partPath in cParts)
                {
                    var pd = ProcessPartByPath(partPath, new ProcessingOptions
                    {
                        Material = "A36 Steel", SaveChanges = false
                    });
                    if (pd != null && pd.Status == ProcessingStatus.Success)
                    {
                        tubeSuccess++;
                        if (!string.IsNullOrEmpty(pd.Cost.OP20_WorkCenter)) tubeWithOP20++;
                        var mapped = PartDataPropertyMap.ToProperties(pd);
                        AssertEquals(Get(mapped, "rbMaterialType"), "1", $"Batch-{Path.GetFileName(partPath)}: rbMaterialType=1");
                    }
                }
                AssertTrue(tubeSuccess > 0, $"Batch: at least 1 C-series success (got {tubeSuccess}/{cParts.Length})");
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
                    Log($"    Result: Status={pd.Status}, Classification={pd.Classification}, OP20_WC={pd.Cost.OP20_WorkCenter ?? "NULL"}, OptiMaterial={pd.OptiMaterial ?? "NULL"}");
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

        private static string Get(Dictionary<string, string> d, string key)
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

        private void Log(string message)
        {
            _log.AppendLine(message);
            ErrorHandler.DebugLog(message);
        }
    }
}
