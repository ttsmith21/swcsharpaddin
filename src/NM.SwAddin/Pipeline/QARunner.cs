using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using NM.Core;
using NM.Core.DataModel;
using NM.SwAddin.AssemblyProcessing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Pipeline
{
    /// <summary>
    /// Runs QA/regression tests by processing all parts in a folder and outputting results JSON.
    /// Triggered by the RunQA toolbar command or programmatically.
    /// </summary>
    public class QARunner
    {
        private readonly ISldWorks _swApp;
        private const string DefaultConfigPath = @"C:\Temp\nm_qa_config.json";

        // File-based debug logger for QA runs
        private static string _qaDebugLogPath;

        public QARunner(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Write a debug message to the QA debug log file.
        /// Uses ErrorHandler.DebugLog which writes to both console and the additional log file.
        /// </summary>
        public static void QALog(string message)
        {
            ErrorHandler.DebugLog(message);
        }

        /// <summary>
        /// Run QA tests using the default config file location.
        /// </summary>
        public QARunSummary Run()
        {
            return Run(DefaultConfigPath);
        }

        /// <summary>
        /// Run QA tests using a specific config file.
        /// </summary>
        public QARunSummary Run(string configPath)
        {
            const string proc = nameof(QARunner) + ".Run";
            ErrorHandler.PushCallStack(proc);
            var summary = new QARunSummary
            {
                RunId = DateTime.Now.ToString("yyyyMMdd_HHmmss"),
                StartedAt = DateTime.UtcNow
            };

            try
            {
                // 1. Read config
                var config = ReadConfig(configPath);
                if (config == null)
                {
                    ErrorHandler.HandleError(proc, $"Failed to read config from {configPath}");
                    return summary;
                }

                if (string.IsNullOrWhiteSpace(config.InputPath) || !Directory.Exists(config.InputPath))
                {
                    ErrorHandler.HandleError(proc, $"Invalid input path: {config.InputPath}");
                    return summary;
                }

                // Initialize QA debug log in output directory
                if (!string.IsNullOrWhiteSpace(config.OutputPath))
                {
                    var outputDir = Path.GetDirectoryName(config.OutputPath);
                    if (!string.IsNullOrEmpty(outputDir))
                    {
                        if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);
                        _qaDebugLogPath = Path.Combine(outputDir, "debug.log");
                        // Set the additional log path in ErrorHandler so all DebugLog calls get captured
                        ErrorHandler.AdditionalDebugLogPath = _qaDebugLogPath;
                        // Clear previous log and write header
                        try { File.WriteAllText(_qaDebugLogPath, $"=== QA Debug Log Started {DateTime.Now:yyyy-MM-dd HH:mm:ss} ==={System.Environment.NewLine}"); }
                        catch { }
                    }
                }

                QALog($"[{proc}] Config loaded: InputPath={config.InputPath}, OutputPath={config.OutputPath}");

                // 2. Collect all part files and assembly files
                var partFiles = Directory.GetFiles(config.InputPath, "*.sldprt", SearchOption.TopDirectoryOnly);
                var assyFiles = Directory.GetFiles(config.InputPath, "*.sldasm", SearchOption.TopDirectoryOnly);
                summary.TotalFiles = partFiles.Length + assyFiles.Length;

                QALog($"[{proc}] Found {partFiles.Length} part files and {assyFiles.Length} assembly files in {config.InputPath}");

                // 3. Process each file
                // IMPORTANT: Disable saving to preserve golden input files
                var options = new ProcessingOptions
                {
                    SaveChanges = false
                };
                var sw = Stopwatch.StartNew();

                foreach (var filePath in partFiles)
                {
                    var result = ProcessSingleFile(filePath, options);
                    summary.Results.Add(result);

                    switch (result.Status)
                    {
                        case "Success":
                            summary.Passed++;
                            break;
                        case "Failed":
                            summary.Failed++;
                            break;
                        default:
                            summary.Errors++;
                            break;
                    }
                }

                // 4. Process assembly files
                foreach (var filePath in assyFiles)
                {
                    var result = ProcessAssemblyFile(filePath);
                    summary.Results.Add(result);

                    switch (result.Status)
                    {
                        case "Success":
                            summary.Passed++;
                            break;
                        case "Failed":
                            summary.Failed++;
                            break;
                        default:
                            summary.Errors++;
                            break;
                    }
                }

                sw.Stop();
                summary.TotalElapsedMs = sw.Elapsed.TotalMilliseconds;
                summary.CompletedAt = DateTime.UtcNow;

                // 4a. Export timing data and populate summary
                ExportTimingData(summary, config.OutputPath);

                // 4b. Check for timing regressions against baseline
                CheckForRegressions(summary, config.OutputPath);

                // 4c. Write results
                if (!string.IsNullOrWhiteSpace(config.OutputPath))
                {
                    WriteResults(summary, config.OutputPath);
                }

                QALog($"");
                QALog($"========== QA Run Complete ==========");
                QALog($"[{proc}] Passed: {summary.Passed}, Failed: {summary.Failed}, Errors: {summary.Errors}");
                QALog($"[{proc}] Debug log saved to: {_qaDebugLogPath}");

                // Clear the additional log path to stop capturing
                ErrorHandler.AdditionalDebugLogPath = null;
                return summary;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "QA run failed", ex, ErrorHandler.LogLevel.Error);
                summary.CompletedAt = DateTime.UtcNow;
                return summary;
            }
            finally
            {
                // Always clear the additional log path when done
                ErrorHandler.AdditionalDebugLogPath = null;
                ErrorHandler.PopCallStack();
            }
        }

        private QATestResult ProcessSingleFile(string filePath, ProcessingOptions options)
        {
            const string proc = nameof(QARunner) + ".ProcessSingleFile";
            var fileName = Path.GetFileName(filePath);
            var sw = Stopwatch.StartNew();

            try
            {
                QALog($"");
                QALog($"========== Processing: {fileName} ==========");

                // Open the file (check if already open first, then open if needed)
                var fileOps = new SolidWorksFileOperations(_swApp);
                var doc = fileOps.GetOpenDocumentByPath(filePath);
                if (doc == null)
                {
                    doc = fileOps.OpenSWDocument(filePath, silent: true, readOnly: false);
                }
                if (doc == null)
                {
                    return new QATestResult
                    {
                        FileName = fileName,
                        FilePath = filePath,
                        Status = "Error",
                        Message = "Failed to open file",
                        ElapsedMs = sw.Elapsed.TotalMilliseconds
                    };
                }

                try
                {
                    // Debug: Check IsTube property right after file opens (before any processing)
                    string cfg = "";
                    try { cfg = doc.ConfigurationManager?.ActiveConfiguration?.Name ?? ""; } catch { }
                    string isTubeOnOpen = SolidWorksApiWrapper.GetCustomPropertyValue(doc, "IsTube", cfg);
                    string isTubeGlobal = SolidWorksApiWrapper.GetCustomPropertyValue(doc, "IsTube", "");
                    QALog($"[{proc}] ON OPEN: IsTube(config='{cfg}')='{isTubeOnOpen}', IsTube(global)='{isTubeGlobal}'");

                    // Process using MainRunner
                    var partData = MainRunner.RunSinglePartData(_swApp, doc, options);
                    sw.Stop();

                    // Convert to QATestResult
                    var result = QATestResult.FromPartData(partData, sw.Elapsed.TotalMilliseconds);
                    QALog($"[{proc}] Result: Status={result.Status}, Classification={result.Classification}, Message={result.Message ?? "OK"}");
                    return result;
                }
                finally
                {
                    // Close the document (don't save changes during QA)
                    CloseDocument(doc);
                }
            }
            catch (Exception ex)
            {
                sw.Stop();
                ErrorHandler.HandleError(proc, $"Exception processing {fileName}", ex);
                return new QATestResult
                {
                    FileName = fileName,
                    FilePath = filePath,
                    Status = "Error",
                    Message = $"Exception: {ex.Message}",
                    ElapsedMs = sw.Elapsed.TotalMilliseconds
                };
            }
        }

        private QATestResult ProcessAssemblyFile(string filePath)
        {
            const string proc = nameof(QARunner) + ".ProcessAssemblyFile";
            var fileName = Path.GetFileName(filePath);
            var sw = Stopwatch.StartNew();

            try
            {
                QALog($"");
                QALog($"========== Processing Assembly: {fileName} ==========");

                // Open the assembly file
                int errors = 0, warnings = 0;
                var doc = _swApp.OpenDoc6(filePath,
                    (int)swDocumentTypes_e.swDocASSEMBLY,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                    "", ref errors, ref warnings);

                if (doc == null)
                {
                    return new QATestResult
                    {
                        FileName = fileName,
                        FilePath = filePath,
                        Status = "Error",
                        Message = $"Failed to open assembly (errors={errors}, warnings={warnings})",
                        IsAssembly = true,
                        Classification = "Assembly",
                        ElapsedMs = sw.Elapsed.TotalMilliseconds
                    };
                }

                try
                {
                    var asm = doc as IAssemblyDoc;
                    if (asm == null)
                    {
                        return new QATestResult
                        {
                            FileName = fileName,
                            FilePath = filePath,
                            Status = "Error",
                            Message = "Document is not an assembly",
                            IsAssembly = true,
                            Classification = "Assembly",
                            ElapsedMs = sw.Elapsed.TotalMilliseconds
                        };
                    }

                    // Use AssemblyComponentQuantifier to collect component info
                    var quantifier = new AssemblyComponentQuantifier();
                    var components = quantifier.CollectViaRecursion(asm);

                    // Count unique parts and sub-assemblies
                    int uniquePartCount = 0;
                    int subAssemblyCount = 0;
                    int totalComponentCount = 0;

                    foreach (var kvp in components)
                    {
                        totalComponentCount += kvp.Value.Quantity;

                        var path = kvp.Value.FilePath ?? "";
                        if (path.EndsWith(".SLDASM", StringComparison.OrdinalIgnoreCase) ||
                            path.EndsWith(".sldasm", StringComparison.OrdinalIgnoreCase))
                        {
                            subAssemblyCount++;
                        }
                        else
                        {
                            uniquePartCount++;
                        }
                    }

                    sw.Stop();

                    var result = new QATestResult
                    {
                        FileName = fileName,
                        FilePath = filePath,
                        Status = "Success",
                        Classification = "Assembly",
                        IsAssembly = true,
                        TotalComponentCount = totalComponentCount,
                        UniquePartCount = uniquePartCount,
                        SubAssemblyCount = subAssemblyCount,
                        ElapsedMs = sw.Elapsed.TotalMilliseconds
                    };

                    QALog($"[{proc}] Result: Status=Success, Components={totalComponentCount}, UniqueParts={uniquePartCount}, SubAssemblies={subAssemblyCount}");
                    return result;
                }
                finally
                {
                    CloseDocument(doc);
                }
            }
            catch (Exception ex)
            {
                sw.Stop();
                ErrorHandler.HandleError(proc, $"Exception processing assembly {fileName}", ex);
                return new QATestResult
                {
                    FileName = fileName,
                    FilePath = filePath,
                    Status = "Error",
                    Message = $"Exception: {ex.Message}",
                    IsAssembly = true,
                    Classification = "Assembly",
                    ElapsedMs = sw.Elapsed.TotalMilliseconds
                };
            }
        }

        private void CloseDocument(IModelDoc2 doc)
        {
            if (doc == null) return;
            try
            {
                // Mark document as not dirty to prevent save prompts
                // This discards any changes made during processing
                try
                {
                    doc.SetSaveFlag();  // Clears the dirty flag
                }
                catch
                {
                    // SetSaveFlag may not be available on all doc types
                }

                var path = doc.GetPathName();
                if (!string.IsNullOrEmpty(path))
                {
                    _swApp.CloseDoc(path);
                }
                else
                {
                    var title = doc.GetTitle();
                    if (!string.IsNullOrEmpty(title))
                    {
                        _swApp.CloseDoc(title);
                    }
                }
            }
            catch
            {
                // Ignore close errors during QA
            }
        }

        private QAConfig ReadConfig(string configPath)
        {
            try
            {
                if (!File.Exists(configPath))
                {
                    ErrorHandler.DebugLog($"[{nameof(ReadConfig)}] Config file not found: {configPath}");
                    return null;
                }

                var json = File.ReadAllText(configPath);
                return SimpleJsonParser.Deserialize<QAConfig>(json);
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(nameof(ReadConfig), $"Failed to parse config: {configPath}", ex);
                return null;
            }
        }

        /// <summary>
        /// Exports timing data from PerformanceTracker and populates the summary's TimingSummary.
        /// </summary>
        private void ExportTimingData(QARunSummary summary, string outputPath)
        {
            const string proc = nameof(ExportTimingData);
            try
            {
                // Populate timing summary from PerformanceTracker
                var timerSummaries = PerformanceTracker.Instance.GetTimingSummaries();
                foreach (var ts in timerSummaries)
                {
                    summary.TimingSummary[ts.Name] = new QATimingSummary
                    {
                        Count = ts.Count,
                        TotalMs = ts.TotalMs,
                        AvgMs = ts.AvgMs,
                        MinMs = ts.MinMs,
                        MaxMs = ts.MaxMs
                    };
                }

                // Export detailed CSV if output path is provided
                if (!string.IsNullOrWhiteSpace(outputPath))
                {
                    var outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir))
                    {
                        var timingCsvPath = Path.Combine(outputDir, "timing.csv");
                        if (PerformanceTracker.Instance.IsEnabled)
                        {
                            PerformanceTracker.Instance.ExportToCsv(timingCsvPath);
                            QALog($"[{proc}] Timing data exported to: {timingCsvPath}");

                            // Log summary to debug output
                            QALog("");
                            QALog("=== Performance Timing Summary ===");
                            foreach (var ts in timerSummaries)
                            {
                                QALog($"  {ts.Name,-35} Count={ts.Count,3} Total={ts.TotalMs,8:F1}ms Avg={ts.AvgMs,8:F1}ms");
                            }
                        }
                        else
                        {
                            QALog($"[{proc}] Performance tracking disabled - no timing.csv generated");
                        }
                    }
                }

                // Clear timers for next run
                PerformanceTracker.Instance.ClearAllTimers();
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Failed to export timing data", ex, ErrorHandler.LogLevel.Warning);
            }
        }

        /// <summary>
        /// Checks for timing regressions against baseline.
        /// Reports warnings if any operation is >20% slower than baseline.
        /// </summary>
        private void CheckForRegressions(QARunSummary summary, string outputPath)
        {
            const string proc = nameof(CheckForRegressions);
            try
            {
                if (string.IsNullOrWhiteSpace(outputPath)) return;

                var outputDir = Path.GetDirectoryName(outputPath);
                if (string.IsNullOrEmpty(outputDir)) return;

                // Look for baseline in tests/ directory (one level up from output run folder)
                var testsDir = Path.GetDirectoryName(outputDir);
                if (string.IsNullOrEmpty(testsDir)) return;

                var baselinePath = Path.Combine(testsDir, "timing-baseline.json");
                if (!File.Exists(baselinePath))
                {
                    QALog($"[{proc}] No baseline found at {baselinePath} - skipping regression check");
                    QALog($"[{proc}] To create a baseline, copy a good timing.csv to timing-baseline.json format");
                    return;
                }

                // Read baseline
                var baselineJson = File.ReadAllText(baselinePath);
                var baseline = SimpleBaselineParser.Parse(baselineJson);
                if (baseline == null || baseline.Timers == null || baseline.Timers.Count == 0)
                {
                    QALog($"[{proc}] Could not parse baseline or baseline has no timers");
                    return;
                }

                // Detect regressions
                var detector = new TimingRegressionDetector();
                var regressions = detector.DetectRegressions(summary.TimingSummary, baseline);

                if (regressions.Count > 0)
                {
                    QALog("");
                    QALog("=== TIMING REGRESSIONS DETECTED ===");
                    foreach (var r in regressions)
                    {
                        QALog($"  {r}");
                    }
                    QALog($"Baseline from: {baseline.CreatedAt}");
                }
                else
                {
                    QALog($"[{proc}] No regressions detected (baseline from {baseline.CreatedAt})");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Regression check failed", ex, ErrorHandler.LogLevel.Warning);
            }
        }

        private void WriteResults(QARunSummary summary, string outputPath)
        {
            try
            {
                // Ensure directory exists
                var dir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var json = SimpleJsonWriter.Serialize(summary);
                File.WriteAllText(outputPath, json);
                ErrorHandler.DebugLog($"[{nameof(WriteResults)}] Results written to: {outputPath}");
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(nameof(WriteResults), $"Failed to write results to {outputPath}", ex);
            }
        }
    }

    /// <summary>
    /// Simple JSON parser for reading config files without external dependencies.
    /// Only supports basic flat object deserialization for QAConfig.
    /// </summary>
    internal static class SimpleJsonParser
    {
        public static T Deserialize<T>(string json) where T : class, new()
        {
            if (typeof(T) != typeof(QAConfig))
            {
                throw new NotSupportedException("SimpleJsonParser only supports QAConfig");
            }

            var config = new QAConfig();

            // Simple regex-based parsing for {"key": "value"} pairs
            var pattern = @"""(\w+)""\s*:\s*""([^""\\]*(?:\\.[^""\\]*)*)""";
            var matches = Regex.Matches(json, pattern);

            foreach (Match m in matches)
            {
                var key = m.Groups[1].Value;
                var value = m.Groups[2].Value
                    .Replace("\\\\", "\x00")  // Temp placeholder for escaped backslash
                    .Replace("\\\"", "\"")
                    .Replace("\\n", "\n")
                    .Replace("\\r", "\r")
                    .Replace("\\t", "\t")
                    .Replace("\x00", "\\");

                switch (key)
                {
                    case "InputPath":
                        config.InputPath = value;
                        break;
                    case "OutputPath":
                        config.OutputPath = value;
                        break;
                    case "BaselinePath":
                        config.BaselinePath = value;
                        break;
                }
            }

            return config as T;
        }
    }

    /// <summary>
    /// Simple parser for timing baseline JSON files.
    /// </summary>
    internal static class SimpleBaselineParser
    {
        public static TimingBaseline Parse(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return null;

            var baseline = new TimingBaseline();

            // Parse top-level string fields
            var stringPattern = @"""(\w+)""\s*:\s*""([^""\\]*(?:\\.[^""\\]*)*)""";
            var stringMatches = Regex.Matches(json, stringPattern);
            foreach (Match m in stringMatches)
            {
                var key = m.Groups[1].Value;
                var value = m.Groups[2].Value;
                switch (key)
                {
                    case "CreatedAt": baseline.CreatedAt = value; break;
                    case "RunId": baseline.RunId = value; break;
                }
            }

            // Parse top-level numeric fields
            var numPattern = @"""(\w+)""\s*:\s*(-?\d+\.?\d*)";
            var numMatches = Regex.Matches(json, numPattern);
            foreach (Match m in numMatches)
            {
                var key = m.Groups[1].Value;
                var value = m.Groups[2].Value;
                switch (key)
                {
                    case "FileCount":
                        if (int.TryParse(value, out var fc)) baseline.FileCount = fc;
                        break;
                    case "TotalElapsedMs":
                        if (double.TryParse(value, out var te)) baseline.TotalElapsedMs = te;
                        break;
                }
            }

            // Parse Timers section - look for nested objects
            // Pattern: "TimerName": { "AvgMs": 123.4, "MaxMs": 456.7, "Count": 10 }
            var timerPattern = @"""(\w+)""\s*:\s*\{\s*""AvgMs""\s*:\s*(-?\d+\.?\d*)\s*,\s*""MaxMs""\s*:\s*(-?\d+\.?\d*)\s*,\s*""Count""\s*:\s*(\d+)\s*\}";
            var timerMatches = Regex.Matches(json, timerPattern);
            foreach (Match m in timerMatches)
            {
                var name = m.Groups[1].Value;
                var entry = new BaselineTimerEntry();
                if (double.TryParse(m.Groups[2].Value, out var avg)) entry.AvgMs = avg;
                if (double.TryParse(m.Groups[3].Value, out var max)) entry.MaxMs = max;
                if (int.TryParse(m.Groups[4].Value, out var count)) entry.Count = count;
                baseline.Timers[name] = entry;
            }

            return baseline;
        }
    }

    /// <summary>
    /// Simple JSON writer for QA results without external dependencies.
    /// </summary>
    internal static class SimpleJsonWriter
    {
        public static string Serialize(QARunSummary summary)
        {
            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine($"  \"RunId\": \"{Escape(summary.RunId)}\",");
            sb.AppendLine($"  \"StartedAt\": \"{summary.StartedAt:O}\",");
            sb.AppendLine($"  \"CompletedAt\": \"{summary.CompletedAt:O}\",");
            sb.AppendLine($"  \"TotalFiles\": {summary.TotalFiles},");
            sb.AppendLine($"  \"Passed\": {summary.Passed},");
            sb.AppendLine($"  \"Failed\": {summary.Failed},");
            sb.AppendLine($"  \"Errors\": {summary.Errors},");
            sb.AppendLine($"  \"TotalElapsedMs\": {summary.TotalElapsedMs:F1},");
            sb.AppendLine("  \"Results\": [");

            for (int i = 0; i < summary.Results.Count; i++)
            {
                var r = summary.Results[i];
                sb.AppendLine("    {");
                sb.AppendLine($"      \"FileName\": \"{Escape(r.FileName)}\",");
                sb.AppendLine($"      \"FilePath\": \"{Escape(r.FilePath)}\",");
                sb.AppendLine($"      \"Status\": \"{Escape(r.Status)}\",");

                if (!string.IsNullOrEmpty(r.Message))
                    sb.AppendLine($"      \"Message\": \"{Escape(r.Message)}\",");

                sb.AppendLine($"      \"Classification\": \"{Escape(r.Classification)}\",");

                // Geometry - only include if non-null
                if (r.Thickness_in.HasValue)
                    sb.AppendLine($"      \"Thickness_in\": {r.Thickness_in.Value:F6},");
                if (r.Mass_lb.HasValue)
                    sb.AppendLine($"      \"Mass_lb\": {r.Mass_lb.Value:F4},");
                if (r.BBoxLength_in.HasValue)
                    sb.AppendLine($"      \"BBoxLength_in\": {r.BBoxLength_in.Value:F4},");
                if (r.BBoxWidth_in.HasValue)
                    sb.AppendLine($"      \"BBoxWidth_in\": {r.BBoxWidth_in.Value:F4},");
                if (r.BBoxHeight_in.HasValue)
                    sb.AppendLine($"      \"BBoxHeight_in\": {r.BBoxHeight_in.Value:F4},");

                // Sheet metal
                if (r.BendCount.HasValue)
                    sb.AppendLine($"      \"BendCount\": {r.BendCount.Value},");
                if (r.BendsBothDirections.HasValue)
                    sb.AppendLine($"      \"BendsBothDirections\": {r.BendsBothDirections.Value.ToString().ToLower()},");
                if (r.FlatArea_sqin.HasValue)
                    sb.AppendLine($"      \"FlatArea_sqin\": {r.FlatArea_sqin.Value:F4},");
                if (r.CutLength_in.HasValue)
                    sb.AppendLine($"      \"CutLength_in\": {r.CutLength_in.Value:F4},");

                // Tube
                if (r.TubeOD_in.HasValue)
                    sb.AppendLine($"      \"TubeOD_in\": {r.TubeOD_in.Value:F4},");
                if (r.TubeWall_in.HasValue)
                    sb.AppendLine($"      \"TubeWall_in\": {r.TubeWall_in.Value:F4},");
                if (r.TubeID_in.HasValue)
                    sb.AppendLine($"      \"TubeID_in\": {r.TubeID_in.Value:F4},");
                if (r.TubeLength_in.HasValue)
                    sb.AppendLine($"      \"TubeLength_in\": {r.TubeLength_in.Value:F4},");
                if (!string.IsNullOrEmpty(r.TubeNPS))
                    sb.AppendLine($"      \"TubeNPS\": \"{Escape(r.TubeNPS)}\",");
                if (!string.IsNullOrEmpty(r.TubeSchedule))
                    sb.AppendLine($"      \"TubeSchedule\": \"{Escape(r.TubeSchedule)}\",");

                // Assembly
                if (r.IsAssembly.HasValue)
                    sb.AppendLine($"      \"IsAssembly\": {r.IsAssembly.Value.ToString().ToLower()},");
                if (r.TotalComponentCount.HasValue)
                    sb.AppendLine($"      \"TotalComponentCount\": {r.TotalComponentCount.Value},");
                if (r.UniquePartCount.HasValue)
                    sb.AppendLine($"      \"UniquePartCount\": {r.UniquePartCount.Value},");
                if (r.SubAssemblyCount.HasValue)
                    sb.AppendLine($"      \"SubAssemblyCount\": {r.SubAssemblyCount.Value},");

                // Material
                if (!string.IsNullOrEmpty(r.Material))
                    sb.AppendLine($"      \"Material\": \"{Escape(r.Material)}\",");
                if (!string.IsNullOrEmpty(r.MaterialCategory))
                    sb.AppendLine($"      \"MaterialCategory\": \"{Escape(r.MaterialCategory)}\",");

                // Costing
                if (r.MaterialCost.HasValue)
                    sb.AppendLine($"      \"MaterialCost\": {r.MaterialCost.Value:F2},");
                if (r.LaserCost.HasValue)
                    sb.AppendLine($"      \"LaserCost\": {r.LaserCost.Value:F2},");
                if (r.BendCost.HasValue)
                    sb.AppendLine($"      \"BendCost\": {r.BendCost.Value:F2},");
                if (r.TapCost.HasValue)
                    sb.AppendLine($"      \"TapCost\": {r.TapCost.Value:F2},");
                if (r.DeburCost.HasValue)
                    sb.AppendLine($"      \"DeburCost\": {r.DeburCost.Value:F2},");
                if (r.TotalCost.HasValue)
                    sb.AppendLine($"      \"TotalCost\": {r.TotalCost.Value:F2},");

                sb.AppendLine($"      \"ElapsedMs\": {r.ElapsedMs:F1},");
                sb.AppendLine($"      \"ProcessedAt\": \"{r.ProcessedAt:O}\"");

                sb.Append("    }");
                if (i < summary.Results.Count - 1)
                    sb.AppendLine(",");
                else
                    sb.AppendLine();
            }

            sb.AppendLine("  ],");

            // TimingSummary section
            sb.AppendLine("  \"TimingSummary\": {");
            int tsIndex = 0;
            foreach (var kvp in summary.TimingSummary)
            {
                tsIndex++;
                sb.AppendLine($"    \"{Escape(kvp.Key)}\": {{");
                sb.AppendLine($"      \"Count\": {kvp.Value.Count},");
                sb.AppendLine($"      \"TotalMs\": {kvp.Value.TotalMs:F1},");
                sb.AppendLine($"      \"AvgMs\": {kvp.Value.AvgMs:F1},");
                sb.AppendLine($"      \"MinMs\": {kvp.Value.MinMs:F1},");
                sb.AppendLine($"      \"MaxMs\": {kvp.Value.MaxMs:F1}");
                sb.Append("    }");
                if (tsIndex < summary.TimingSummary.Count)
                    sb.AppendLine(",");
                else
                    sb.AppendLine();
            }
            sb.AppendLine("  }");
            sb.AppendLine("}");

            return sb.ToString();
        }

        private static string Escape(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n")
                .Replace("\t", "\\t");
        }
    }
}
