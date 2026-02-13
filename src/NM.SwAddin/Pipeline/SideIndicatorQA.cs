using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NM.Core;
using NM.SwAddin.SheetMetal;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Pipeline
{
    /// <summary>
    /// Automated QA test for the Side Indicator (green/red face coloring) toggle.
    /// Validates the 3-state cycle: Baseline → Applied → Restored.
    /// Verifies that original face colors are properly saved and restored.
    /// </summary>
    public class SideIndicatorQA
    {
        private readonly StringBuilder _log = new StringBuilder();
        private int _passed;
        private int _failed;

        /// <summary>
        /// Runs the Side Indicator QA test suite.
        /// </summary>
        /// <param name="swApp">SolidWorks application instance</param>
        /// <param name="inputDir">Path to GoldStandard_Inputs directory</param>
        /// <returns>0 = all passed, 1 = failures</returns>
        public int Run(ISldWorks swApp, string inputDir)
        {
            Log("=== SideIndicatorQA START ===");
            Log($"Input: {inputDir}");
            Log("");

            // Test 1: Native sheet metal bracket (has FlatPattern feature)
            string b1 = Path.Combine(inputDir, "B1_NativeBracket_14ga_CS.SLDPRT");
            if (File.Exists(b1))
                RunToggleTest(swApp, b1, "B1_NativeBracket (native SM, FlatPattern)");
            else
                Log($"SKIP: {b1} not found");

            Log("");

            // Test 2: Imported sheet metal bracket (fallback to largest face normal)
            string b2 = Path.Combine(inputDir, "B2_ImportedBracket_14ga_CS.SLDPRT");
            if (File.Exists(b2))
                RunToggleTest(swApp, b2, "B2_ImportedBracket (imported SM, fallback normal)");
            else
                Log($"SKIP: {b2} not found");

            Log("");

            // Test 3: Idempotency — toggle ON/OFF/ON/OFF on native bracket
            if (File.Exists(b1))
                RunIdempotencyTest(swApp, b1, "B1 Idempotency (2 full cycles)");
            else
                Log($"SKIP: idempotency test — {b1} not found");

            Log("");

            // Summary
            int total = _passed + _failed;
            Log("=== SideIndicatorQA SUMMARY ===");
            Log($"Total assertions: {total}");
            Log($"Passed: {_passed}");
            Log($"Failed: {_failed}");
            Log($"Result: {(_failed == 0 ? "ALL PASSED" : "FAILURES DETECTED")}");

            // Write log
            string outputDir = Path.Combine(Path.GetDirectoryName(inputDir) ?? inputDir, "Run_SideIndicatorQA");
            try
            {
                if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);
                File.WriteAllText(Path.Combine(outputDir, "side_indicator_qa.log"), _log.ToString());
                Console.WriteLine($"Log written to: {Path.Combine(outputDir, "side_indicator_qa.log")}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Warning: could not write log: {ex.Message}");
            }

            Console.WriteLine(_log.ToString());

            return _failed == 0 ? 0 : 1;
        }

        /// <summary>
        /// Core 3-state toggle test:
        /// STATE 1 (Baseline) → Toggle ON → STATE 2 (Applied) → Toggle OFF → STATE 3 (Restored)
        /// Verifies that STATE 3 matches STATE 1.
        /// </summary>
        private void RunToggleTest(ISldWorks swApp, string filePath, string testName)
        {
            Log($"--- TEST: {testName} ---");
            Log($"File: {filePath}");

            IModelDoc2 model = null;
            try
            {
                model = OpenPart(swApp, filePath);
                if (model == null)
                {
                    AssertFail("OpenPart", $"Could not open {Path.GetFileName(filePath)}");
                    return;
                }

                var service = new SideIndicatorService();

                // STATE 1: Baseline — read all face colors before any toggle
                var baseline = ReadAllFaceColors(model);
                Log($"  STATE 1 (Baseline): {baseline.Count} faces read");
                LogColorSummary("    Baseline", baseline);
                AssertTrue("Baseline has faces", baseline.Count > 0);

                // ACTION 1: Toggle ON
                service.Toggle(swApp);
                AssertTrue("Service reports active after toggle ON", service.IsActive(model));

                // STATE 2: Applied — read all face colors after toggle ON
                var applied = ReadAllFaceColors(model);
                Log($"  STATE 2 (Applied): {applied.Count} faces read");
                LogColorSummary("    Applied", applied);
                AssertTrue("Applied has faces", applied.Count > 0);

                // Verify green, red, and gray faces are present
                int greenCount = CountFacesWithColor(applied, 0.0, 0.8, 0.0);
                int redCount = CountFacesWithColor(applied, 0.8, 0.0, 0.0);
                int grayCount = CountFacesWithColor(applied, 0.7, 0.7, 0.7);
                Log($"  Applied colors: green={greenCount}, red={redCount}, gray={grayCount}");
                AssertTrue("At least 1 green (top) face", greenCount >= 1);
                AssertTrue("At least 1 red (bottom) face", redCount >= 1);
                AssertTrue("At least 1 gray (edge/bend) face", grayCount >= 1);

                // Verify colors changed from baseline
                int changedCount = CountChangedFaces(baseline, applied);
                AssertTrue("All faces changed from baseline", changedCount == baseline.Count);
                Log($"  Changed from baseline: {changedCount}/{baseline.Count}");

                // ACTION 2: Toggle OFF
                service.Toggle(swApp);
                AssertTrue("Service reports inactive after toggle OFF", !service.IsActive(model));

                // STATE 3: Restored — read all face colors after toggle OFF
                var restored = ReadAllFaceColors(model);
                Log($"  STATE 3 (Restored): {restored.Count} faces read");
                LogColorSummary("    Restored", restored);

                // Verify restored matches baseline
                AssertTrue("Same face count after restore", restored.Count == baseline.Count);
                int matchCount = CountMatchingFaces(baseline, restored);
                Log($"  Restored matches baseline: {matchCount}/{baseline.Count}");
                AssertTrue("All faces restored to baseline", matchCount == baseline.Count);

                // Log any mismatches for debugging
                if (matchCount < baseline.Count)
                {
                    LogMismatches(baseline, restored);
                }

                Log($"  TEST {testName}: {(_failed == 0 ? "PASSED" : "HAS FAILURES")}");
            }
            catch (Exception ex)
            {
                AssertFail("Unexpected exception", ex.Message);
                Log($"  Stack: {ex.StackTrace}");
            }
            finally
            {
                if (model != null)
                    ClosePart(swApp, model);
            }
        }

        /// <summary>
        /// Idempotency test: toggle ON/OFF twice and verify restoration works both times.
        /// </summary>
        private void RunIdempotencyTest(ISldWorks swApp, string filePath, string testName)
        {
            Log($"--- TEST: {testName} ---");
            Log($"File: {filePath}");

            IModelDoc2 model = null;
            int prevFailed = _failed;
            try
            {
                model = OpenPart(swApp, filePath);
                if (model == null)
                {
                    AssertFail("OpenPart", $"Could not open {Path.GetFileName(filePath)}");
                    return;
                }

                var service = new SideIndicatorService();
                var baseline = ReadAllFaceColors(model);
                Log($"  Baseline: {baseline.Count} faces");

                for (int cycle = 1; cycle <= 2; cycle++)
                {
                    Log($"  Cycle {cycle}: Toggle ON...");
                    service.Toggle(swApp);
                    AssertTrue($"Cycle {cycle}: active after ON", service.IsActive(model));

                    var applied = ReadAllFaceColors(model);
                    int greenCount = CountFacesWithColor(applied, 0.0, 0.8, 0.0);
                    int redCount = CountFacesWithColor(applied, 0.8, 0.0, 0.0);
                    Log($"    Applied: green={greenCount}, red={redCount}");
                    AssertTrue($"Cycle {cycle}: has green faces", greenCount >= 1);

                    Log($"  Cycle {cycle}: Toggle OFF...");
                    service.Toggle(swApp);
                    AssertTrue($"Cycle {cycle}: inactive after OFF", !service.IsActive(model));

                    var restored = ReadAllFaceColors(model);
                    int matchCount = CountMatchingFaces(baseline, restored);
                    Log($"    Restored: {matchCount}/{baseline.Count} match baseline");
                    AssertTrue($"Cycle {cycle}: all faces restored", matchCount == baseline.Count);
                }

                Log($"  TEST {testName}: {(_failed == prevFailed ? "PASSED" : "HAS FAILURES")}");
            }
            catch (Exception ex)
            {
                AssertFail("Unexpected exception", ex.Message);
            }
            finally
            {
                if (model != null)
                    ClosePart(swApp, model);
            }
        }

        #region Face Color Reading

        /// <summary>
        /// Represents a face's color state. Color is null if no face-level override exists.
        /// </summary>
        private class FaceColorEntry
        {
            public int Index;
            public double[] Color; // null = no face-level override (inheriting from body/part)
        }

        /// <summary>
        /// Reads MaterialPropertyValues for every face on every solid body in the model.
        /// Returns a list indexed by face order (consistent for same model state).
        /// </summary>
        private List<FaceColorEntry> ReadAllFaceColors(IModelDoc2 model)
        {
            var result = new List<FaceColorEntry>();
            var part = model as IPartDoc;
            if (part == null) return result;

            var bodiesRaw = part.GetBodies2((int)swBodyType_e.swSolidBody, true);
            if (bodiesRaw == null) return result;

            int idx = 0;
            foreach (var bodyObj in (object[])bodiesRaw)
            {
                var body = bodyObj as IBody2;
                if (body == null) continue;

                var facesRaw = body.GetFaces();
                if (facesRaw == null) continue;

                foreach (var faceObj in (object[])facesRaw)
                {
                    var face = faceObj as IFace2;
                    if (face == null) continue;

                    double[] color = null;
                    try
                    {
                        var raw = face.MaterialPropertyValues;
                        if (raw is double[] arr && arr.Length >= 3)
                            color = (double[])arr.Clone();
                    }
                    catch { }

                    result.Add(new FaceColorEntry { Index = idx, Color = color });
                    idx++;
                }
            }

            return result;
        }

        #endregion

        #region Color Comparison Utilities

        /// <summary>
        /// Compares two color arrays with tolerance. Both null = match.
        /// </summary>
        private static bool ColorsMatch(double[] a, double[] b, double tolerance = 0.01)
        {
            if (a == null && b == null) return true;
            if (a == null || b == null) return false;
            if (a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++)
            {
                if (Math.Abs(a[i] - b[i]) > tolerance) return false;
            }
            return true;
        }

        private int CountFacesWithColor(List<FaceColorEntry> faces, double r, double g, double b)
        {
            int count = 0;
            foreach (var f in faces)
            {
                if (f.Color != null && f.Color.Length >= 3 &&
                    Math.Abs(f.Color[0] - r) < 0.05 &&
                    Math.Abs(f.Color[1] - g) < 0.05 &&
                    Math.Abs(f.Color[2] - b) < 0.05)
                    count++;
            }
            return count;
        }

        private int CountChangedFaces(List<FaceColorEntry> baseline, List<FaceColorEntry> current)
        {
            int count = 0;
            int limit = Math.Min(baseline.Count, current.Count);
            for (int i = 0; i < limit; i++)
            {
                if (!ColorsMatch(baseline[i].Color, current[i].Color))
                    count++;
            }
            return count;
        }

        private int CountMatchingFaces(List<FaceColorEntry> expected, List<FaceColorEntry> actual)
        {
            int count = 0;
            int limit = Math.Min(expected.Count, actual.Count);
            for (int i = 0; i < limit; i++)
            {
                if (ColorsMatch(expected[i].Color, actual[i].Color))
                    count++;
            }
            return count;
        }

        #endregion

        #region Logging and Assertions

        private void LogColorSummary(string prefix, List<FaceColorEntry> faces)
        {
            int withOverride = faces.Count(f => f.Color != null);
            int inherited = faces.Count(f => f.Color == null);
            Log($"{prefix}: {withOverride} with face-level color, {inherited} inherited (null)");
        }

        private void LogMismatches(List<FaceColorEntry> baseline, List<FaceColorEntry> restored)
        {
            int limit = Math.Min(baseline.Count, restored.Count);
            for (int i = 0; i < limit; i++)
            {
                if (!ColorsMatch(baseline[i].Color, restored[i].Color))
                {
                    string baseStr = FormatColor(baseline[i].Color);
                    string restStr = FormatColor(restored[i].Color);
                    Log($"    MISMATCH face[{i}]: baseline={baseStr}, restored={restStr}");
                }
            }
        }

        private static string FormatColor(double[] c)
        {
            if (c == null) return "null (inherited)";
            if (c.Length >= 3)
                return $"RGB({c[0]:F2},{c[1]:F2},{c[2]:F2})";
            return $"[{string.Join(",", c.Select(v => v.ToString("F2")))}]";
        }

        private void Log(string message)
        {
            _log.AppendLine(message);
        }

        private void AssertTrue(string name, bool condition)
        {
            if (condition)
            {
                _passed++;
                Log($"  PASS: {name}");
            }
            else
            {
                _failed++;
                Log($"  FAIL: {name}");
            }
        }

        private void AssertFail(string name, string reason)
        {
            _failed++;
            Log($"  FAIL: {name} — {reason}");
        }

        #endregion

        #region File Operations

        private IModelDoc2 OpenPart(ISldWorks swApp, string filePath)
        {
            int errors = 0, warnings = 0;
            var doc = swApp.OpenDoc6(filePath,
                (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "", ref errors, ref warnings);

            if (doc == null)
            {
                Log($"  ERROR: OpenDoc6 failed for {Path.GetFileName(filePath)} (errors={errors}, warnings={warnings})");
                return null;
            }

            return doc as IModelDoc2;
        }

        private void ClosePart(ISldWorks swApp, IModelDoc2 model)
        {
            try
            {
                string name = model.GetTitle();
                if (!string.IsNullOrEmpty(name))
                    swApp.CloseDoc(name);
            }
            catch (Exception ex)
            {
                Log($"  Warning: close failed: {ex.Message}");
            }
        }

        #endregion
    }
}
