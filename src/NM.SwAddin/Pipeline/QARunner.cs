using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using NM.Core;
using NM.Core.DataModel;
using NM.Core.Export;
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

                // 2b. ReadPropertiesOnly mode: dump VBA custom properties without processing
                if (config.IsReadPropertiesOnly)
                {
                    QALog($"[{proc}] Mode: ReadPropertiesOnly - extracting custom properties only");
                    var propsMap = new Dictionary<string, Dictionary<string, string>>();
                    var swProps = Stopwatch.StartNew();

                    foreach (var filePath in partFiles)
                    {
                        var props = ReadPropertiesFromFile(filePath);
                        if (props != null)
                            propsMap[Path.GetFileName(filePath)] = props;
                    }
                    foreach (var filePath in assyFiles)
                    {
                        var props = ReadPropertiesFromFile(filePath);
                        if (props != null)
                            propsMap[Path.GetFileName(filePath)] = props;
                    }

                    swProps.Stop();
                    summary.TotalElapsedMs = swProps.Elapsed.TotalMilliseconds;
                    summary.CompletedAt = DateTime.UtcNow;
                    summary.Passed = propsMap.Count;
                    summary.TotalFiles = partFiles.Length + assyFiles.Length;

                    // Write vba_properties.json
                    if (!string.IsNullOrWhiteSpace(config.OutputPath))
                    {
                        var outputDir = Path.GetDirectoryName(config.OutputPath);
                        if (!string.IsNullOrEmpty(outputDir))
                        {
                            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);
                            var propsPath = Path.Combine(outputDir, "vba_properties.json");
                            WritePropertiesJson(propsMap, propsPath);
                            QALog($"[{proc}] Wrote vba_properties.json with {propsMap.Count} files to {propsPath}");
                        }
                    }

                    QALog($"[{proc}] ReadPropertiesOnly complete: {propsMap.Count} files in {swProps.Elapsed.TotalSeconds:F1}s");
                    ErrorHandler.AdditionalDebugLogPath = null;
                    return summary;
                }

                // 3. Process each file
                // IMPORTANT: Disable saving to preserve golden input files
                var options = new ProcessingOptions
                {
                    SaveChanges = false
                };
                var sw = Stopwatch.StartNew();

                foreach (var filePath in partFiles)
                {
                    var result = ProcessSingleFile(filePath, options, summary.PartDataCollection);
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

                // 4a. Stamp BomQty from assembly component quantities onto part results
                foreach (var assyResult in summary.Results)
                {
                    if (assyResult.IsAssembly != true || assyResult.ComponentQuantities == null) continue;
                    foreach (var partResult in summary.Results)
                    {
                        if (partResult.IsAssembly == true) continue;
                        if (assyResult.ComponentQuantities.TryGetValue(partResult.FileName, out int qty))
                            partResult.BomQty = qty;
                    }
                }

                // 4b. Export timing data and populate summary
                ExportTimingData(summary, config.OutputPath);

                // 4c. Check for timing regressions against baseline
                CheckForRegressions(summary, config.OutputPath);

                // 4d. Write results
                if (!string.IsNullOrWhiteSpace(config.OutputPath))
                {
                    WriteResults(summary, config.OutputPath);
                }

                // 4e. Generate C# Import.prn for comparison with VBA baseline
                if (summary.PartDataCollection.Count > 0 && !string.IsNullOrWhiteSpace(config.OutputPath))
                {
                    try
                    {
                        var outputDir = Path.GetDirectoryName(config.OutputPath);
                        var exportData = ErpExportDataBuilder.FromPartDataCollection(
                            summary.PartDataCollection, "GOLD_STANDARD_ASM_CLEAN", "", "GOLD STANDARD TEST");
                        var exporter = new ErpExportFormat();
                        var prnPath = Path.Combine(outputDir, "Import_CSharp.prn");
                        exporter.ExportToImportPrn(exportData, prnPath);
                        QALog($"[{proc}] Generated Import_CSharp.prn at {prnPath}");
                    }
                    catch (Exception prnEx)
                    {
                        QALog($"[{proc}] WARNING: Import.prn generation failed: {prnEx.Message}");
                    }
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

        private QATestResult ProcessSingleFile(string filePath, ProcessingOptions options, List<PartData> partDataOut = null)
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

                    // Retain PartData for ERP export generation
                    if (partDataOut != null && partData != null)
                        partDataOut.Add(partData);

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

                    // Build per-component quantity map for BomQty stamping
                    var compQty = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    foreach (var kvp in components)
                    {
                        var compFileName = System.IO.Path.GetFileName(kvp.Value.FilePath);
                        if (!string.IsNullOrEmpty(compFileName))
                            compQty[compFileName] = kvp.Value.Quantity;
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
                        ComponentQuantities = compQty,
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

        /// <summary>
        /// Opens a file and reads ALL custom properties without running the processing pipeline.
        /// Returns a dictionary of property name -> value, or null on failure.
        /// </summary>
        private Dictionary<string, string> ReadPropertiesFromFile(string filePath)
        {
            const string proc = nameof(QARunner) + ".ReadPropertiesFromFile";
            var fileName = Path.GetFileName(filePath);

            try
            {
                QALog($"[{proc}] Reading properties from: {fileName}");

                // Determine document type
                var ext = Path.GetExtension(filePath).ToUpperInvariant();
                int docType = ext == ".SLDASM" ? (int)swDocumentTypes_e.swDocASSEMBLY : (int)swDocumentTypes_e.swDocPART;

                // Open the file
                int errors = 0, warnings = 0;
                var doc = _swApp.OpenDoc6(filePath, docType,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent | (int)swOpenDocOptions_e.swOpenDocOptions_ReadOnly,
                    "", ref errors, ref warnings);

                if (doc == null)
                {
                    QALog($"[{proc}] Failed to open: {fileName} (errors={errors})");
                    return null;
                }

                try
                {
                    var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    // Read global (file-level) properties
                    string[] names, values;
                    int[] types;
                    if (SolidWorksApiWrapper.GetCustomProperties(doc, "", out names, out types, out values))
                    {
                        for (int i = 0; i < names.Length; i++)
                        {
                            if (!string.IsNullOrEmpty(names[i]))
                                props[names[i]] = values[i] ?? "";
                        }
                    }

                    // Also read config-level properties (may override globals)
                    string cfg = "";
                    try { cfg = doc.ConfigurationManager?.ActiveConfiguration?.Name ?? ""; } catch { }
                    if (!string.IsNullOrEmpty(cfg))
                    {
                        if (SolidWorksApiWrapper.GetCustomProperties(doc, cfg, out names, out types, out values))
                        {
                            for (int i = 0; i < names.Length; i++)
                            {
                                if (!string.IsNullOrEmpty(names[i]))
                                    props[names[i]] = values[i] ?? "";
                            }
                        }
                    }

                    QALog($"[{proc}] Read {props.Count} properties from {fileName}");
                    return props;
                }
                finally
                {
                    CloseDocument(doc);
                }
            }
            catch (Exception ex)
            {
                QALog($"[{proc}] Exception reading {fileName}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Writes the properties map to a JSON file using SimpleJsonWriter pattern.
        /// </summary>
        private static void WritePropertiesJson(Dictionary<string, Dictionary<string, string>> propsMap, string outputPath)
        {
            var sb = new StringBuilder();
            sb.AppendLine("{");
            sb.AppendLine("  \"files\": {");

            int fileIndex = 0;
            foreach (var kvp in propsMap)
            {
                var fileName = kvp.Key;
                var props = kvp.Value;
                sb.AppendLine($"    \"{SimpleJsonWriter.EscapeString(fileName)}\": {{");

                int propIndex = 0;
                foreach (var prop in props)
                {
                    var comma = (propIndex < props.Count - 1) ? "," : "";
                    sb.AppendLine($"      \"{SimpleJsonWriter.EscapeString(prop.Key)}\": \"{SimpleJsonWriter.EscapeString(prop.Value)}\"{comma}");
                    propIndex++;
                }

                var fileComma = (fileIndex < propsMap.Count - 1) ? "," : "";
                sb.AppendLine($"    }}{fileComma}");
                fileIndex++;
            }

            sb.AppendLine("  }");
            sb.AppendLine("}");

            File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);
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
                    case "Mode":
                        config.Mode = value;
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

                if (!string.IsNullOrEmpty(r.Configuration))
                    sb.AppendLine($"      \"Configuration\": \"{Escape(r.Configuration)}\",");

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

                // BOM quantity (from assembly traversal)
                if (r.BomQty.HasValue)
                    sb.AppendLine($"      \"BomQty\": {r.BomQty.Value},");

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
                if (r.RollCost.HasValue)
                    sb.AppendLine($"      \"RollCost\": {r.RollCost.Value:F2},");
                if (r.TotalCost.HasValue)
                    sb.AppendLine($"      \"TotalCost\": {r.TotalCost.Value:F2},");

                // Cost breakdown dictionary
                if (r.CostBreakdown != null && r.CostBreakdown.Count > 0)
                {
                    sb.AppendLine("      \"CostBreakdown\": {");
                    int cbIdx = 0;
                    foreach (var cb in r.CostBreakdown)
                    {
                        cbIdx++;
                        sb.Append($"        \"{Escape(cb.Key)}\": {cb.Value:F2}");
                        sb.AppendLine(cbIdx < r.CostBreakdown.Count ? "," : "");
                    }
                    sb.AppendLine("      },");
                }

                // Per-workcenter setup/run times (hours)
                if (r.F115_Setup.HasValue)
                    sb.AppendLine($"      \"F115_Setup\": {r.F115_Setup.Value:F4},");
                if (r.F115_Run.HasValue)
                    sb.AppendLine($"      \"F115_Run\": {r.F115_Run.Value:F4},");
                if (r.F140_Setup.HasValue)
                    sb.AppendLine($"      \"F140_Setup\": {r.F140_Setup.Value:F4},");
                if (r.F140_Run.HasValue)
                    sb.AppendLine($"      \"F140_Run\": {r.F140_Run.Value:F4},");
                if (r.F210_Setup.HasValue)
                    sb.AppendLine($"      \"F210_Setup\": {r.F210_Setup.Value:F4},");
                if (r.F210_Run.HasValue)
                    sb.AppendLine($"      \"F210_Run\": {r.F210_Run.Value:F4},");
                if (r.F220_Setup.HasValue)
                    sb.AppendLine($"      \"F220_Setup\": {r.F220_Setup.Value:F4},");
                if (r.F220_Run.HasValue)
                    sb.AppendLine($"      \"F220_Run\": {r.F220_Run.Value:F4},");
                if (r.F325_Setup.HasValue)
                    sb.AppendLine($"      \"F325_Setup\": {r.F325_Setup.Value:F4},");
                if (r.F325_Run.HasValue)
                    sb.AppendLine($"      \"F325_Run\": {r.F325_Run.Value:F4},");

                // Tube laser routing (F110)
                if (r.F110_Setup.HasValue)
                    sb.AppendLine($"      \"F110_Setup\": {r.F110_Setup.Value:F4},");
                if (r.F110_Run.HasValue)
                    sb.AppendLine($"      \"F110_Run\": {r.F110_Run.Value:F4},");
                // 5-axis laser routing (N145)
                if (r.N145_Setup.HasValue)
                    sb.AppendLine($"      \"N145_Setup\": {r.N145_Setup.Value:F4},");
                if (r.N145_Run.HasValue)
                    sb.AppendLine($"      \"N145_Run\": {r.N145_Run.Value:F4},");
                // Saw routing (F300)
                if (r.F300_Setup.HasValue)
                    sb.AppendLine($"      \"F300_Setup\": {r.F300_Setup.Value:F4},");
                if (r.F300_Run.HasValue)
                    sb.AppendLine($"      \"F300_Run\": {r.F300_Run.Value:F4},");
                // Purchased part routing (NPUR)
                if (r.NPUR_Setup.HasValue)
                    sb.AppendLine($"      \"NPUR_Setup\": {r.NPUR_Setup.Value:F4},");
                if (r.NPUR_Run.HasValue)
                    sb.AppendLine($"      \"NPUR_Run\": {r.NPUR_Run.Value:F4},");
                // Customer-supplied routing (CUST)
                if (r.CUST_Setup.HasValue)
                    sb.AppendLine($"      \"CUST_Setup\": {r.CUST_Setup.Value:F4},");
                if (r.CUST_Run.HasValue)
                    sb.AppendLine($"      \"CUST_Run\": {r.CUST_Run.Value:F4},");

                // ERP fields
                if (!string.IsNullOrEmpty(r.OptiMaterial))
                    sb.AppendLine($"      \"OptiMaterial\": \"{Escape(r.OptiMaterial)}\",");
                if (!string.IsNullOrEmpty(r.Description))
                    sb.AppendLine($"      \"Description\": \"{Escape(r.Description)}\",");

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
            return EscapeString(value);
        }

        internal static string EscapeString(string value)
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
