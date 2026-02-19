using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NM.Core;
using NM.Core.Drawing;
using NM.SwAddin.Drawing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Pipeline
{
    /// <summary>
    /// Automated QA test for drawing creation and dimensioning.
    /// Tests the full drawing pipeline: create drawing → drop view → dimension → validate.
    /// Processes B-series (sheet metal) and C-series (tube) gold standard parts.
    /// </summary>
    public class DrawingQA
    {
        private readonly StringBuilder _log = new StringBuilder();
        private int _passed;
        private int _failed;
        private readonly List<DrawingTestResult> _results = new List<DrawingTestResult>();

        /// <summary>
        /// Per-part test result for JSON output.
        /// </summary>
        private class DrawingTestResult
        {
            public string FileName;
            public string Status; // "Pass", "Fail", "Error", "Skip"
            public string Message;
            public bool DrawingCreated;
            public string DrawingPath;
            public string DxfPath;
            public int ViewCount;
            public int DimensionCount;
            public int BendLinesFound;
            public int DanglingDimensions;
            public int UndimensionedBendLines;
            public int ViewsOnSheet;
            public int ViewsOffSheet;
            public double SheetWidth_in;
            public double SheetHeight_in;
            public double ElapsedMs;
            public List<string> Warnings = new List<string>();
            public List<ViewPositionInfo> ViewPositions = new List<ViewPositionInfo>();
        }

        private class ViewPositionInfo
        {
            public string Name;
            public double XMin_in, YMin_in, XMax_in, YMax_in;
            public double Width_in, Height_in;
            public bool FullyOnSheet;
            public bool PartiallyOffSheet;
            public bool FullyOffSheet;
        }

        /// <summary>
        /// Runs the Drawing QA test suite.
        /// </summary>
        /// <param name="swApp">SolidWorks application instance</param>
        /// <param name="inputDir">Path to GoldStandard_Inputs directory</param>
        /// <returns>0 = all passed, 1 = failures</returns>
        public int Run(ISldWorks swApp, string inputDir)
        {
            _swApp = swApp;

            Log("=== DrawingQA START ===");
            Log($"Input: {inputDir}");
            Log($"Time:  {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            Log("");

            // Create temp output directory for drawings (don't pollute test inputs)
            string outputDir = Path.Combine(
                Path.GetDirectoryName(inputDir) ?? inputDir,
                $"Run_DrawingQA_{DateTime.Now:yyyyMMdd_HHmmss}");
            Directory.CreateDirectory(outputDir);
            Log($"Output: {outputDir}");
            Log("");

            // Test B-series: Sheet metal parts (should get flat pattern view + dimensions)
            var sheetMetalParts = new[]
            {
                new { File = "B1_NativeBracket_14ga_CS.SLDPRT", Desc = "Native SM bracket" },
                new { File = "B2_ImportedBracket_14ga_CS.SLDPRT", Desc = "Imported SM bracket" },
                new { File = "B3_RolledCylinder_16ga_SS.SLDPRT", Desc = "Rolled cylinder" },
                new { File = "B4_SheetMetal_TappedHoles.SLDPRT", Desc = "SM with tapped holes" },
                new { File = "B5_SheetMetal_TappedHoles_SS.SLDPRT", Desc = "SM tapped holes SS" },
                new { File = "B6_SheetMetal_NamedTapHoles.SLDPRT", Desc = "SM named tap holes" },
            };

            Log("--- SHEET METAL PARTS ---");
            foreach (var part in sheetMetalParts)
            {
                string path = Path.Combine(inputDir, part.File);
                if (File.Exists(path))
                    RunDrawingTest(swApp, path, part.Desc, outputDir, isSheetMetal: true);
                else
                    LogSkip(part.File, "File not found");
            }

            Log("");

            // Test C-series: Tube/structural parts (should get standard views + dimensions)
            var tubeParts = new[]
            {
                new { File = "C1_RoundTube_2OD_SCH40.SLDPRT", Desc = "Round tube" },
                new { File = "C2_RectTube_2x1.SLDPRT", Desc = "Rectangular tube" },
                new { File = "C3_SquareTube_2x2.SLDPRT", Desc = "Square tube" },
            };

            Log("--- TUBE PARTS ---");
            foreach (var part in tubeParts)
            {
                string path = Path.Combine(inputDir, part.File);
                if (File.Exists(path))
                    RunDrawingTest(swApp, path, part.Desc, outputDir, isSheetMetal: false);
                else
                    LogSkip(part.File, "File not found");
            }

            Log("");

            // Summary
            int totalTests = _results.Count;
            int passCount = _results.Count(r => r.Status == "Pass");
            int failCount = _results.Count(r => r.Status == "Fail");
            int errorCount = _results.Count(r => r.Status == "Error");
            int skipCount = _results.Count(r => r.Status == "Skip");

            Log("=== DrawingQA SUMMARY ===");
            Log($"Total parts tested: {totalTests}");
            Log($"Passed:  {passCount}");
            Log($"Failed:  {failCount}");
            Log($"Errors:  {errorCount}");
            Log($"Skipped: {skipCount}");
            Log($"Assertions: {_passed} passed, {_failed} failed");
            Log($"Result: {(_failed == 0 && failCount == 0 && errorCount == 0 ? "ALL PASSED" : "FAILURES DETECTED")}");

            // Write log file
            try
            {
                File.WriteAllText(Path.Combine(outputDir, "drawing_qa.log"), _log.ToString());
                Console.WriteLine($"Log: {Path.Combine(outputDir, "drawing_qa.log")}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Warning: could not write log: {ex.Message}");
            }

            // Write JSON results
            try
            {
                string json = SerializeResults(outputDir);
                File.WriteAllText(Path.Combine(outputDir, "drawing_results.json"), json);
                Console.WriteLine($"Results: {Path.Combine(outputDir, "drawing_results.json")}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Warning: could not write results JSON: {ex.Message}");
            }

            Console.WriteLine(_log.ToString());

            return (_failed == 0 && failCount == 0 && errorCount == 0) ? 0 : 1;
        }

        /// <summary>
        /// Core test: open part → process (so OP20 is set) → create drawing → validate.
        /// </summary>
        private void RunDrawingTest(ISldWorks swApp, string filePath, string testName,
            string outputDir, bool isSheetMetal)
        {
            string fileName = Path.GetFileName(filePath);
            Log($"--- TEST: {testName} ({fileName}) ---");

            var result = new DrawingTestResult { FileName = fileName };
            var sw = System.Diagnostics.Stopwatch.StartNew();

            IModelDoc2 model = null;
            IModelDoc2 drawModel = null;
            try
            {
                // Step 1: Open the part
                model = OpenPart(swApp, filePath);
                if (model == null)
                {
                    result.Status = "Error";
                    result.Message = "Could not open part";
                    AssertFail("OpenPart", $"Could not open {fileName}");
                    _results.Add(result);
                    return;
                }

                // Step 2: Ensure OP20 is set (required by DrawingGenerator).
                // We set it directly rather than running the full pipeline, because:
                //   a) The pipeline shows a "complete" message box (bad for headless QA)
                //   b) Pipeline testing is handled by --qa; drawing QA should be isolated
                //   c) We know the part type from the test matrix
                string existingOp20 = GetCustomProperty(model, "OP20");
                if (string.IsNullOrEmpty(existingOp20))
                {
                    string op20Value = isSheetMetal ? "F115 FLAT LASER" : "F300 SAW";
                    // Set at both file-level and config-level (DrawingGenerator checks config first)
                    SetCustomProperty(model, "", "OP20", op20Value);
                    string configName = model.ConfigurationManager?.ActiveConfiguration?.Name ?? "";
                    if (!string.IsNullOrEmpty(configName))
                        SetCustomProperty(model, configName, "OP20", op20Value);
                    Log($"  Set OP20: {op20Value}");
                }
                else
                {
                    Log($"  OP20 already set: {existingOp20}");
                }

                // Step 3: Create the drawing
                Log($"  Creating drawing...");
                var generator = new DrawingGenerator(swApp);
                var drawOptions = new DrawingGenerator.DrawingOptions
                {
                    OutputFolder = outputDir,
                    SaveDrawing = true,
                    CreateDxf = true,
                    IncludeFlatPattern = isSheetMetal,
                    IncludeFormedView = isSheetMetal,
                    IncludeDimensions = true,
                };

                var drawResult = generator.CreateDrawing(model, drawOptions);

                result.DrawingCreated = drawResult.Success;
                result.DrawingPath = drawResult.DrawingPath;
                result.DxfPath = drawResult.DxfPath;
                result.DimensionCount = drawResult.DimensionsAdded;

                AssertTrue($"Drawing created for {fileName}", drawResult.Success);
                Log($"  Result: {drawResult.Message}");
                Log($"  Dimensions added by generator: {drawResult.DimensionsAdded}");

                if (!drawResult.Success)
                {
                    result.Status = "Fail";
                    result.Message = drawResult.Message;
                    _results.Add(result);
                    ClosePart(swApp, model);
                    return;
                }

                // Step 4: Validate the drawing
                Log($"  Validating drawing...");

                // Find the drawing document (it should be the active doc now, or we open it)
                drawModel = swApp.ActiveDoc as IModelDoc2;
                if (drawModel == null || (swDocumentTypes_e)drawModel.GetType() != swDocumentTypes_e.swDocDRAWING)
                {
                    // Try to activate the drawing
                    if (!string.IsNullOrEmpty(drawResult.DrawingPath))
                    {
                        int err = 0;
                        swApp.ActivateDoc3(
                            Path.GetFileName(drawResult.DrawingPath), false, 0, ref err);
                        drawModel = swApp.ActiveDoc as IModelDoc2;
                    }
                }

                if (drawModel != null && (swDocumentTypes_e)drawModel.GetType() == swDocumentTypes_e.swDocDRAWING)
                {
                    var drawDoc = drawModel as IDrawingDoc;
                    if (drawDoc != null)
                    {
                        ValidateDrawing(drawDoc, drawModel, result, isSheetMetal);
                    }
                    else
                    {
                        result.Warnings.Add("Could not cast to IDrawingDoc");
                    }
                }
                else
                {
                    result.Warnings.Add("Drawing document not active after creation");
                    Log($"  WARNING: Could not access drawing document for validation");
                }

                // Step 5: Validate files on disk
                if (!string.IsNullOrEmpty(drawResult.DrawingPath))
                    AssertTrue($"Drawing file exists: {Path.GetFileName(drawResult.DrawingPath)}",
                        File.Exists(drawResult.DrawingPath));

                if (!string.IsNullOrEmpty(drawResult.DxfPath))
                    AssertTrue($"DXF file exists: {Path.GetFileName(drawResult.DxfPath)}",
                        File.Exists(drawResult.DxfPath));

                result.Status = "Pass";
                result.Message = "Drawing created and validated";
                Log($"  TEST {testName}: PASSED");
            }
            catch (Exception ex)
            {
                result.Status = "Error";
                result.Message = ex.Message;
                AssertFail("Unexpected exception", ex.Message);
                Log($"  Stack: {ex.StackTrace}");
            }
            finally
            {
                sw.Stop();
                result.ElapsedMs = sw.Elapsed.TotalMilliseconds;
                _results.Add(result);

                // Close drawing first, then part
                if (drawModel != null)
                    CloseDoc(swApp, drawModel);
                if (model != null)
                    ClosePart(swApp, model);

                Log($"  Elapsed: {result.ElapsedMs:F0}ms");
                Log("");
            }
        }

        /// <summary>
        /// Validates the created drawing: view count, view positions relative to sheet,
        /// dimension count, bend lines, dangling dims.
        /// </summary>
        private void ValidateDrawing(IDrawingDoc drawDoc, IModelDoc2 drawModel,
            DrawingTestResult result, bool isSheetMetal)
        {
            const double MetersToInches = 39.3700787401575;

            // Get sheet size
            var sheet = drawDoc.GetCurrentSheet() as ISheet;
            if (sheet == null)
            {
                result.Warnings.Add("No current sheet");
                return;
            }

            var sheetProps = sheet.GetProperties2() as double[];
            double sheetW = 0, sheetH = 0;
            if (sheetProps != null && sheetProps.Length >= 8)
            {
                // GetProperties2 returns: [0]=PaperSize, [1]=?, [2]=Scale1, [3]=Scale2,
                // [4]=FirstAngle, [5]=Width(m), [6]=Height(m), [7]=?
                sheetW = sheetProps[5]; // meters
                sheetH = sheetProps[6]; // meters
            }

            // Fallback: try GetSize
            if (sheetW <= 0 || sheetH <= 0)
            {
                double w = 0, h = 0;
                sheet.GetSize(ref w, ref h);
                sheetW = w;
                sheetH = h;
            }

            double sheetW_in = sheetW * MetersToInches;
            double sheetH_in = sheetH * MetersToInches;
            result.SheetWidth_in = Math.Round(sheetW_in, 2);
            result.SheetHeight_in = Math.Round(sheetH_in, 2);
            Log($"  Sheet size: {sheetW_in:F2}\" x {sheetH_in:F2}\"");

            // Get all views
            var viewsRaw = sheet.GetViews() as object[];
            int viewCount = viewsRaw?.Length ?? 0;
            result.ViewCount = viewCount;
            Log($"  Views on sheet: {viewCount}");
            // Sheet metal: flat pattern + 0-2 projected views (depends on bend directions)
            // Tube: end view + side view
            if (isSheetMetal)
                AssertTrue("Sheet metal: at least 1 view (flat pattern)", viewCount >= 1);
            else
                AssertTrue("Tube: at least 2 views (end + side)", viewCount >= 2);

            // Check each view's position relative to the sheet
            int onSheet = 0, offSheet = 0;
            int totalDims = 0;
            IView flatView = null;
            bool foundSecondaryView = false;

            if (viewsRaw != null)
            {
                // Track primary view name (first view is typically Drawing View1)
                string primaryViewName = null;

                foreach (var viewObj in viewsRaw)
                {
                    var view = viewObj as IView;
                    if (view == null) continue;

                    string viewName = view.GetName2() ?? "(unnamed)";

                    // Track primary view
                    if (primaryViewName == null)
                        primaryViewName = viewName;

                    // Track flat pattern view for later validation
                    if (isSheetMetal && flatView == null &&
                        viewName.IndexOf("Flat", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        flatView = view;
                    }

                    // Detect secondary views
                    if (viewName != primaryViewName)
                    {
                        foundSecondaryView = true;
                        string config = view.ReferencedConfiguration ?? "";
                        Log($"    Secondary view '{viewName}' config='{config}'");

                        // For sheet metal projected views, verify they use Default config
                        if (isSheetMetal && config.Equals("Default", StringComparison.OrdinalIgnoreCase))
                        {
                            Log($"    Projected view '{viewName}' correctly uses Default config");
                        }
                    }

                    // Get view bounding box (meters, in sheet coordinates)
                    var outline = view.GetOutline() as double[];
                    if (outline != null && outline.Length >= 4)
                    {
                        double xMin = outline[0] * MetersToInches;
                        double yMin = outline[1] * MetersToInches;
                        double xMax = outline[2] * MetersToInches;
                        double yMax = outline[3] * MetersToInches;
                        double vw = xMax - xMin;
                        double vh = yMax - yMin;

                        var vp = new ViewPositionInfo
                        {
                            Name = viewName,
                            XMin_in = Math.Round(xMin, 3),
                            YMin_in = Math.Round(yMin, 3),
                            XMax_in = Math.Round(xMax, 3),
                            YMax_in = Math.Round(yMax, 3),
                            Width_in = Math.Round(vw, 3),
                            Height_in = Math.Round(vh, 3),
                        };

                        // Check if view is within sheet bounds (with small tolerance for rounding)
                        const double tol = 0.05; // 0.05" tolerance
                        bool leftOk = xMin >= -tol;
                        bool bottomOk = yMin >= -tol;
                        bool rightOk = xMax <= sheetW_in + tol;
                        bool topOk = yMax <= sheetH_in + tol;

                        if (leftOk && bottomOk && rightOk && topOk)
                        {
                            vp.FullyOnSheet = true;
                            onSheet++;
                        }
                        else
                        {
                            // Check if completely off sheet vs partially off
                            bool completelyOff = xMax < -tol || xMin > sheetW_in + tol ||
                                                 yMax < -tol || yMin > sheetH_in + tol;
                            if (completelyOff)
                            {
                                vp.FullyOffSheet = true;
                            }
                            else
                            {
                                vp.PartiallyOffSheet = true;
                            }
                            offSheet++;

                            string issue = completelyOff ? "FULLY OFF SHEET" : "PARTIALLY OFF SHEET";
                            Log($"    WARNING: {viewName} {issue}");
                            Log($"      View bounds: ({xMin:F2}\", {yMin:F2}\") to ({xMax:F2}\", {yMax:F2}\")");
                            Log($"      Sheet bounds: (0, 0) to ({sheetW_in:F2}\", {sheetH_in:F2}\")");
                        }

                        result.ViewPositions.Add(vp);
                        Log($"    {viewName}: ({xMin:F2}\",{yMin:F2}\") to ({xMax:F2}\",{yMax:F2}\") [{vw:F2}\"x{vh:F2}\"] {(vp.FullyOnSheet ? "OK" : "OFF")}");
                    }
                    else
                    {
                        Log($"    {viewName}: no outline data");
                    }

                    // Count dimensions in this view
                    var annots = view.GetAnnotations() as object[];
                    if (annots != null)
                    {
                        foreach (var annotObj in annots)
                        {
                            var annot = annotObj as IAnnotation;
                            if (annot == null) continue;
                            if (annot.GetType() == (int)swAnnotationType_e.swDisplayDimension)
                                totalDims++;
                        }
                    }
                }
            }

            result.ViewsOnSheet = onSheet;
            result.ViewsOffSheet = offSheet;
            // Preserve any dimension count already set by generator (use max of generator count and view annotation count)
            if (totalDims > result.DimensionCount)
                result.DimensionCount = totalDims;

            Log($"  Views on sheet: {onSheet}, off sheet: {offSheet}");
            Log($"  Dimensions found in views: {totalDims}");

            // Assert that all views are on the sheet
            AssertTrue("All views within sheet bounds", offSheet == 0);

            // Assert secondary view was created (required for tubes; optional for sheet metal
            // since projected views only appear when bends exist in that direction)
            if (!isSheetMetal)
                AssertTrue("Secondary view present (tube side view)", foundSecondaryView);

            // Assert dimensions were added (Phase 3)
            AssertTrue("At least 1 dimension added", result.DimensionCount > 0);

            // Run the DrawingValidator
            try
            {
                var validator = new DrawingValidator(_swApp);
                var validation = validator.Validate(drawDoc, flatView);

                result.DanglingDimensions = validation.DanglingDimensions.Count;
                result.UndimensionedBendLines = validation.UndimensionedBendLines.Count;

                if (validation.DanglingDimensions.Count > 0)
                {
                    Log($"  Dangling dimensions: {validation.DanglingDimensions.Count}");
                    foreach (var d in validation.DanglingDimensions)
                        Log($"    - {d}");
                }

                if (validation.UndimensionedBendLines.Count > 0)
                {
                    Log($"  Undimensioned bend lines: {validation.UndimensionedBendLines.Count}");
                    foreach (var bl in validation.UndimensionedBendLines)
                        Log($"    - {bl}");
                }

                foreach (var w in validation.Warnings)
                {
                    result.Warnings.Add(w);
                    Log($"  Validation warning: {w}");
                }
            }
            catch (Exception ex)
            {
                result.Warnings.Add($"Validation exception: {ex.Message}");
                Log($"  Validation error: {ex.Message}");
            }
        }

        /// <summary>
        /// ISldWorks field for validator construction (set during Run).
        /// </summary>
        private ISldWorks _swApp;

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
            CloseDoc(swApp, model);
        }

        private void CloseDoc(ISldWorks swApp, IModelDoc2 model)
        {
            try
            {
                // Use GetPathName first (more reliable than GetTitle for COM)
                string path = model.GetPathName();
                string name = !string.IsNullOrEmpty(path) ? Path.GetFileName(path) : model.GetTitle();
                if (!string.IsNullOrEmpty(name))
                    swApp.CloseDoc(name);
            }
            catch
            {
                // RPC_E_DISCONNECTED is expected if SW already cleaned up the doc
            }
        }

        private string GetCustomProperty(IModelDoc2 model, string propName)
        {
            try
            {
                string configName = model.ConfigurationManager?.ActiveConfiguration?.Name ?? "";
                var propMgr = model.Extension?.CustomPropertyManager[configName];
                if (propMgr == null)
                    propMgr = model.Extension?.CustomPropertyManager[""];

                if (propMgr != null)
                {
                    string valOut = "";
                    string resolvedOut = "";
                    bool wasResolved = false;
                    propMgr.Get5(propName, true, out valOut, out resolvedOut, out wasResolved);
                    return resolvedOut ?? valOut ?? "";
                }
            }
            catch { }
            return "";
        }

        private void SetCustomProperty(IModelDoc2 model, string configName, string propName, string value)
        {
            try
            {
                var propMgr = model.Extension?.CustomPropertyManager[configName];
                if (propMgr != null)
                {
                    propMgr.Add3(propName,
                        (int)swCustomInfoType_e.swCustomInfoText,
                        value,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                }
            }
            catch { }
        }

        #endregion

        #region Logging and Assertions

        private void Log(string message)
        {
            _log.AppendLine(message);
        }

        private void LogSkip(string fileName, string reason)
        {
            Log($"  SKIP: {fileName} — {reason}");
            _results.Add(new DrawingTestResult
            {
                FileName = fileName,
                Status = "Skip",
                Message = reason,
            });
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

        #region JSON Serialization

        private string SerializeResults(string outputDir)
        {
            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine($"  \"RunId\": \"{DateTime.Now:yyyyMMdd_HHmmss}\",");
            sb.AppendLine($"  \"StartedAt\": \"{DateTime.UtcNow:O}\",");
            sb.AppendLine($"  \"OutputDir\": \"{EscapeJson(outputDir)}\",");
            sb.AppendLine($"  \"TotalParts\": {_results.Count},");
            sb.AppendLine($"  \"Passed\": {_results.Count(r => r.Status == "Pass")},");
            sb.AppendLine($"  \"Failed\": {_results.Count(r => r.Status == "Fail")},");
            sb.AppendLine($"  \"Errors\": {_results.Count(r => r.Status == "Error")},");
            sb.AppendLine($"  \"Skipped\": {_results.Count(r => r.Status == "Skip")},");
            sb.AppendLine($"  \"Assertions\": {{ \"passed\": {_passed}, \"failed\": {_failed} }},");
            sb.AppendLine("  \"Results\": [");

            for (int i = 0; i < _results.Count; i++)
            {
                var r = _results[i];
                sb.AppendLine("    {");
                sb.AppendLine($"      \"FileName\": \"{EscapeJson(r.FileName)}\",");
                sb.AppendLine($"      \"Status\": \"{r.Status}\",");
                sb.AppendLine($"      \"Message\": \"{EscapeJson(r.Message ?? "")}\",");
                sb.AppendLine($"      \"DrawingCreated\": {(r.DrawingCreated ? "true" : "false")},");
                sb.AppendLine($"      \"DrawingPath\": \"{EscapeJson(r.DrawingPath ?? "")}\",");
                sb.AppendLine($"      \"DxfPath\": \"{EscapeJson(r.DxfPath ?? "")}\",");
                sb.AppendLine($"      \"ViewCount\": {r.ViewCount},");
                sb.AppendLine($"      \"DimensionCount\": {r.DimensionCount},");
                sb.AppendLine($"      \"BendLinesFound\": {r.BendLinesFound},");
                sb.AppendLine($"      \"DanglingDimensions\": {r.DanglingDimensions},");
                sb.AppendLine($"      \"UndimensionedBendLines\": {r.UndimensionedBendLines},");
                sb.AppendLine($"      \"SheetSize_in\": \"{r.SheetWidth_in}x{r.SheetHeight_in}\",");
                sb.AppendLine($"      \"ViewsOnSheet\": {r.ViewsOnSheet},");
                sb.AppendLine($"      \"ViewsOffSheet\": {r.ViewsOffSheet},");
                sb.AppendLine($"      \"ElapsedMs\": {r.ElapsedMs:F1},");

                // View positions
                sb.Append("      \"ViewPositions\": [");
                if (r.ViewPositions.Count > 0)
                {
                    sb.AppendLine();
                    for (int v = 0; v < r.ViewPositions.Count; v++)
                    {
                        var vp = r.ViewPositions[v];
                        sb.Append($"        {{ \"Name\": \"{EscapeJson(vp.Name)}\", ");
                        sb.Append($"\"Bounds_in\": [{vp.XMin_in},{vp.YMin_in},{vp.XMax_in},{vp.YMax_in}], ");
                        sb.Append($"\"Size_in\": [{vp.Width_in},{vp.Height_in}], ");
                        sb.Append($"\"OnSheet\": {(vp.FullyOnSheet ? "true" : "false")} }}");
                        if (v < r.ViewPositions.Count - 1) sb.Append(",");
                        sb.AppendLine();
                    }
                    sb.Append("      ");
                }
                sb.AppendLine("],");

                sb.Append("      \"Warnings\": [");
                if (r.Warnings.Count > 0)
                {
                    sb.AppendLine();
                    for (int w = 0; w < r.Warnings.Count; w++)
                    {
                        sb.Append($"        \"{EscapeJson(r.Warnings[w])}\"");
                        if (w < r.Warnings.Count - 1) sb.Append(",");
                        sb.AppendLine();
                    }
                    sb.Append("      ");
                }
                sb.AppendLine("]");

                sb.Append("    }");
                if (i < _results.Count - 1) sb.Append(",");
                sb.AppendLine();
            }

            sb.AppendLine("  ]");
            sb.AppendLine("}");
            return sb.ToString();
        }

        private static string EscapeJson(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n");
        }

        #endregion
    }
}
