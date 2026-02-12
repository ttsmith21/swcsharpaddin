using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NM.Core.Config;
using SolidWorks.Interop.sldworks;

namespace NM.BatchRunner
{
    class Program
    {
        private static readonly string CrashLogPath = @"C:\Temp\nm_batch_crash.txt";

        static int Main(string[] args)
        {
            // Set up global exception handlers for crash logging
            AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
            {
                var ex = e.ExceptionObject as Exception;
                LogCrash("UnhandledException", ex);
            };

            Console.WriteLine("NM.BatchRunner - Headless SolidWorks Automation");
            Console.WriteLine("================================================");
            Console.WriteLine($"Started at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            Console.WriteLine();

            if (args.Length == 0 || args[0] == "--help" || args[0] == "-h")
            {
                PrintUsage();
                return 1;
            }

            // Initialize NM configuration (laser speeds, tube cutting params, etc.)
            // BatchRunner bypasses the add-in lifecycle, so we must init explicitly.
            var repoRoot = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\"));
            var configDir = Path.Combine(repoRoot, "config");
            NmConfigProvider.Initialize(configDir);

            try
            {
                using (var launcher = new SolidWorksLauncher())
                {
                    Console.WriteLine("Starting SolidWorks...");
                    var swApp = launcher.Start();
                    Console.WriteLine("Connected to SolidWorks.");

                    if (args[0] == "--qa")
                    {
                        return RunQA(swApp);
                    }
                    else if (args[0] == "--property-qa")
                    {
                        return RunPropertyQA(swApp);
                    }
                    else if (args[0] == "--step-qa")
                    {
                        return RunStepQA(swApp, args);
                    }
                    else if (args[0] == "--dump")
                    {
                        return RunDump(swApp, args);
                    }
                    else if (args[0] == "--pipeline")
                    {
                        return RunPipeline(swApp, args);
                    }
                    else
                    {
                        Console.Error.WriteLine($"Unknown command: {args[0]}");
                        PrintUsage();
                        return 1;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
                LogCrash("Main", ex);
                return 1;
            }
            finally
            {
                Console.WriteLine();
                Console.WriteLine($"Finished at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            }
        }

        static void LogCrash(string context, Exception ex)
        {
            try
            {
                var logEntry = $"=== CRASH LOG ===\r\n" +
                    $"Time: {DateTime.Now:O}\r\n" +
                    $"Context: {context}\r\n" +
                    $"Exception: {ex?.GetType().Name ?? "Unknown"}\r\n" +
                    $"Message: {ex?.Message ?? "No message"}\r\n" +
                    $"Stack:\r\n{ex?.StackTrace ?? "No stack trace"}\r\n" +
                    $"Inner: {ex?.InnerException?.Message ?? "None"}\r\n" +
                    $"================\r\n\r\n";

                File.AppendAllText(CrashLogPath, logEntry);
                Console.Error.WriteLine($"Crash logged to: {CrashLogPath}");
            }
            catch
            {
                // Can't log, just continue
            }
        }

        static void PrintUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  NM.BatchRunner.exe --property-qa");
            Console.WriteLine("      Run property write QA: exercises single-part, assembly, batch workflows");
            Console.WriteLine("      and validates OP20, OptiMaterial, rbMaterialType are correctly populated.");
            Console.WriteLine();
            Console.WriteLine("  NM.BatchRunner.exe --qa [--config <path>]");
            Console.WriteLine("      Run QA tests on parts specified in config file");
            Console.WriteLine("      Default config: C:\\Temp\\nm_qa_config.json");
            Console.WriteLine();
            Console.WriteLine("  NM.BatchRunner.exe --step-qa [--file <path>]");
            Console.WriteLine("      Import a STEP assembly, externalize components, run pipeline on each part");
            Console.WriteLine("      Default file: tests\\GoldStandard_Inputs - Copy\\Large\\6 hours.step");
            Console.WriteLine();
            Console.WriteLine("  NM.BatchRunner.exe --dump --file <path> [--tag <tag>]");
            Console.WriteLine("      Dump all custom properties from a SolidWorks part to JSON");
            Console.WriteLine();
            Console.WriteLine("  NM.BatchRunner.exe --pipeline --file <path>");
            Console.WriteLine("      Run processing pipeline on a single file");
            Console.WriteLine();
            Console.WriteLine("  NM.BatchRunner.exe --help");
            Console.WriteLine("      Show this help message");
        }

        static int RunPropertyQA(ISldWorks swApp)
        {
            Console.WriteLine("Running Property Write QA...");

            // Find the gold standard inputs directory (relative to exe location)
            var repoRoot = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\"));
            var inputDir = Path.Combine(repoRoot, @"tests\GoldStandard_Inputs");
            var outputDir = Path.Combine(repoRoot, "tests",
                $"Run_PropertyQA_{DateTime.Now:yyyyMMdd_HHmmss}");

            if (!Directory.Exists(inputDir))
            {
                Console.Error.WriteLine($"Error: Gold standard inputs not found: {inputDir}");
                return 1;
            }

            var qa = new NM.SwAddin.Pipeline.PropertyWriteQA(swApp, inputDir, outputDir);
            return qa.Run();
        }

        static int RunQA(ISldWorks swApp)
        {
            Console.WriteLine("Running QA tests...");

            var runner = new NM.SwAddin.Pipeline.QARunner(swApp);
            var summary = runner.Run();

            Console.WriteLine();
            Console.WriteLine("=== QA Results ===");
            Console.WriteLine($"Total:   {summary.TotalFiles}");
            Console.WriteLine($"Passed:  {summary.Passed}");
            Console.WriteLine($"Failed:  {summary.Failed}");
            Console.WriteLine($"Errors:  {summary.Errors}");
            Console.WriteLine($"Time:    {summary.TotalElapsedMs:F0}ms");

            // Write summary for scripts to read
            File.WriteAllText(@"C:\Temp\nm_qa_summary.txt",
                $"QA Complete: Total={summary.TotalFiles}, Passed={summary.Passed}, Failed={summary.Failed}, Errors={summary.Errors}, Time={summary.TotalElapsedMs:F0}ms");

            return (summary.Failed > 0 || summary.Errors > 0) ? 1 : 0;
        }

        static int RunStepQA(ISldWorks swApp, string[] args)
        {
            // Parse optional --file override
            string stepFile = null;
            for (int i = 1; i < args.Length - 1; i++)
            {
                if (args[i] == "--file")
                {
                    stepFile = args[i + 1];
                    break;
                }
            }

            // Default source path
            if (string.IsNullOrEmpty(stepFile))
            {
                var repoRoot = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\"));
                stepFile = Path.Combine(repoRoot, @"tests\GoldStandard_Inputs - Copy\Large\6 hours.step");
            }

            if (!File.Exists(stepFile))
            {
                Console.Error.WriteLine($"Error: STEP file not found: {stepFile}");
                return 1;
            }

            Console.WriteLine($"STEP QA source: {stepFile}");

            // Create timestamped run folder
            var repoRootForOutput = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\"));
            var runFolder = Path.Combine(repoRootForOutput, "tests",
                $"Run_StepQA_{DateTime.Now:yyyyMMdd_HHmmss}");
            Directory.CreateDirectory(runFolder);

            // Copy STEP file into run folder
            var destStep = Path.Combine(runFolder, Path.GetFileName(stepFile));
            File.Copy(stepFile, destStep, overwrite: true);
            Console.WriteLine($"Copied STEP to: {destStep}");
            Console.WriteLine($"Run folder: {runFolder}");

            // Run the STEP QA pipeline
            var runner = new NM.SwAddin.Pipeline.QARunner(swApp);
            var summary = runner.RunStepImport(destStep, runFolder);

            Console.WriteLine();
            Console.WriteLine("=== STEP QA Results ===");
            Console.WriteLine($"Total:   {summary.TotalFiles}");
            Console.WriteLine($"Passed:  {summary.Passed}");
            Console.WriteLine($"Failed:  {summary.Failed}");
            Console.WriteLine($"Errors:  {summary.Errors}");
            Console.WriteLine($"Time:    {summary.TotalElapsedMs:F0}ms");

            // Write summary for scripts to read
            File.WriteAllText(Path.Combine(runFolder, "summary.txt"),
                $"STEP QA Complete: Total={summary.TotalFiles}, Passed={summary.Passed}, Failed={summary.Failed}, Errors={summary.Errors}, Time={summary.TotalElapsedMs:F0}ms");

            return (summary.Failed > 0 || summary.Errors > 0) ? 1 : 0;
        }

        static int RunDump(ISldWorks swApp, string[] args)
        {
            string filePath = null;
            string tag = null;
            for (int i = 1; i < args.Length - 1; i++)
            {
                if (args[i] == "--file") filePath = args[i + 1];
                if (args[i] == "--tag") tag = args[i + 1];
            }

            if (string.IsNullOrEmpty(filePath))
            {
                Console.Error.WriteLine("Error: --dump requires --file <path>");
                return 1;
            }

            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"Error: File not found: {filePath}");
                return 1;
            }

            Console.WriteLine($"Dumping properties from: {filePath}");

            // Open the file (silent, read-only)
            int errors = 0, warnings = 0;
            var doc = swApp.OpenDoc6(filePath,
                (int)SolidWorks.Interop.swconst.swDocumentTypes_e.swDocPART,
                (int)SolidWorks.Interop.swconst.swOpenDocOptions_e.swOpenDocOptions_Silent
                    | (int)SolidWorks.Interop.swconst.swOpenDocOptions_e.swOpenDocOptions_ReadOnly,
                "", ref errors, ref warnings);

            if (doc == null)
            {
                Console.Error.WriteLine($"Error: Could not open file (errors={errors}, warnings={warnings})");
                return 1;
            }

            try
            {
                var result = new System.Collections.Generic.Dictionary<string, object>();
                result["Timestamp"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                result["PartFile"] = filePath;
                result["PartName"] = Path.GetFileNameWithoutExtension(filePath);
                result["Tag"] = tag ?? "";

                // File-level custom properties
                var fileProps = new System.Collections.Generic.Dictionary<string, object>();
                var fileMgr = doc.Extension.get_CustomPropertyManager("");
                var fileNames = (string[])fileMgr.GetNames();
                if (fileNames != null)
                {
                    foreach (var name in fileNames)
                    {
                        string valOut = "", resolvedOut = "";
                        bool wasResolved = false;
                        fileMgr.Get5(name, false, out valOut, out resolvedOut, out wasResolved);
                        fileProps[name] = new Dictionary<string, string> { { "Value", valOut }, { "Resolved", resolvedOut } };
                        Console.WriteLine($"  [File] {name} = {resolvedOut}");
                    }
                }
                result["FileProperties"] = fileProps;

                // Per-configuration custom properties
                var cfgProps = new System.Collections.Generic.Dictionary<string, object>();
                var configNames = (string[])doc.GetConfigurationNames();
                if (configNames != null)
                {
                    foreach (var cfgName in configNames)
                    {
                        var cfg = new System.Collections.Generic.Dictionary<string, object>();
                        var cfgMgr = doc.Extension.get_CustomPropertyManager(cfgName);
                        var cfgPropNames = (string[])cfgMgr.GetNames();
                        if (cfgPropNames != null)
                        {
                            foreach (var name in cfgPropNames)
                            {
                                string valOut = "", resolvedOut = "";
                                bool wasResolved = false;
                                cfgMgr.Get5(name, false, out valOut, out resolvedOut, out wasResolved);
                                cfg[name] = new Dictionary<string, string> { { "Value", valOut }, { "Resolved", resolvedOut } };
                                Console.WriteLine($"  [{cfgName}] {name} = {resolvedOut}");
                            }
                        }
                        cfgProps[cfgName] = cfg;
                    }
                }
                result["ConfigProperties"] = cfgProps;

                // Model info
                var modelInfo = new System.Collections.Generic.Dictionary<string, object>();
                try
                {
                    var massProp = doc.Extension.CreateMassProperty();
                    if (massProp != null)
                    {
                        modelInfo["Mass_kg"] = Math.Round(massProp.Mass, 6);
                        modelInfo["Mass_lb"] = Math.Round(massProp.Mass * 2.20462, 4);
                        modelInfo["Volume_m3"] = Math.Round(massProp.Volume, 9);
                        modelInfo["SurfaceArea_m2"] = Math.Round(massProp.SurfaceArea, 6);
                    }
                }
                catch { }

                try
                {
                    var partDoc = (SolidWorks.Interop.sldworks.PartDoc)doc;
                    var bodies = partDoc.GetBodies2(0, true);
                    modelInfo["SolidBodyCount"] = bodies != null ? ((object[])bodies).Length : 0;
                }
                catch { }

                try
                {
                    var activeCfg = doc.ConfigurationManager.ActiveConfiguration;
                    if (activeCfg != null) modelInfo["ActiveConfig"] = activeCfg.Name;
                }
                catch { }

                result["ModelInfo"] = modelInfo;

                // Serialize to JSON
                var suffix = string.IsNullOrEmpty(tag) ? "" : $"_{tag}";
                var outPath = Path.Combine(Path.GetDirectoryName(filePath),
                    $"{Path.GetFileNameWithoutExtension(filePath)}{suffix}_properties.json");

                var json = SerializePropertyDump(result);
                File.WriteAllText(outPath, json);

                Console.WriteLine();
                Console.WriteLine($"Properties saved to: {outPath}");
                Console.WriteLine($"Total: {fileProps.Count} file props, {cfgProps.Count} config(s)");
                return 0;
            }
            finally
            {
                swApp.CloseDoc(doc.GetTitle());
            }
        }

        static int RunPipeline(ISldWorks swApp, string[] args)
        {
            string filePath = null;
            for (int i = 1; i < args.Length - 1; i++)
            {
                if (args[i] == "--file")
                {
                    filePath = args[i + 1];
                    break;
                }
            }

            if (string.IsNullOrEmpty(filePath))
            {
                Console.Error.WriteLine("Error: --pipeline requires --file <path>");
                return 1;
            }

            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine($"Error: File not found: {filePath}");
                return 1;
            }

            Console.WriteLine($"Running pipeline on: {filePath}");

            // Open the file
            int errors = 0, warnings = 0;
            var doc = swApp.OpenDoc6(filePath,
                (int)SolidWorks.Interop.swconst.swDocumentTypes_e.swDocPART,
                (int)SolidWorks.Interop.swconst.swOpenDocOptions_e.swOpenDocOptions_Silent,
                "", ref errors, ref warnings);

            if (doc == null)
            {
                Console.Error.WriteLine($"Error: Could not open file (errors={errors}, warnings={warnings})");
                return 1;
            }

            try
            {
                var dispatcher = new NM.SwAddin.Pipeline.WorkflowDispatcher(swApp);
                dispatcher.Run();

                Console.WriteLine("Pipeline completed successfully.");
                File.WriteAllText(@"C:\Temp\nm_pipeline_complete.txt", $"Pipeline completed at {DateTime.Now}");
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Pipeline error: {ex.Message}");
                File.WriteAllText(@"C:\Temp\nm_pipeline_error.txt", ex.ToString());
                return 1;
            }
        }

        static string SerializePropertyDump(Dictionary<string, object> data)
        {
            var sb = new StringBuilder();
            WriteJsonValue(sb, data, 0);
            return sb.ToString();
        }

        static void WriteJsonValue(StringBuilder sb, object value, int indent)
        {
            if (value == null)
            {
                sb.Append("null");
            }
            else if (value is string s)
            {
                sb.Append('"');
                sb.Append(s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n"));
                sb.Append('"');
            }
            else if (value is bool b)
            {
                sb.Append(b ? "true" : "false");
            }
            else if (value is int || value is long || value is double || value is float || value is decimal)
            {
                sb.Append(value.ToString());
            }
            else if (value is Dictionary<string, object> dict)
            {
                WriteJsonDict(sb, dict, indent);
            }
            else if (value is Dictionary<string, string> sdict)
            {
                sb.Append("{ ");
                bool first = true;
                foreach (var kv in sdict)
                {
                    if (!first) sb.Append(", ");
                    sb.Append('"').Append(kv.Key).Append("\": ");
                    WriteJsonValue(sb, kv.Value, indent + 1);
                    first = false;
                }
                sb.Append(" }");
            }
            else
            {
                sb.Append('"').Append(value.ToString().Replace("\\", "\\\\").Replace("\"", "\\\"")).Append('"');
            }
        }

        static void WriteJsonDict(StringBuilder sb, Dictionary<string, object> dict, int indent)
        {
            var pad = new string(' ', (indent + 1) * 2);
            var closePad = new string(' ', indent * 2);
            sb.AppendLine("{");
            bool first = true;
            foreach (var kv in dict)
            {
                if (!first) sb.AppendLine(",");
                sb.Append(pad).Append('"').Append(kv.Key).Append("\": ");
                WriteJsonValue(sb, kv.Value, indent + 1);
                first = false;
            }
            sb.AppendLine();
            sb.Append(closePad).Append('}');
        }
    }
}
