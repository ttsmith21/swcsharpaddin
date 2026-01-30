using System;
using System.IO;
using SolidWorks.Interop.sldworks;

namespace NM.BatchRunner
{
    class Program
    {
        static int Main(string[] args)
        {
            Console.WriteLine("NM.BatchRunner - Headless SolidWorks Automation");
            Console.WriteLine("================================================");

            if (args.Length == 0 || args[0] == "--help" || args[0] == "-h")
            {
                PrintUsage();
                return 1;
            }

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
                File.WriteAllText(@"C:\Temp\nm_batch_error.txt", ex.ToString());
                return 1;
            }
        }

        static void PrintUsage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  NM.BatchRunner.exe --qa [--config <path>]");
            Console.WriteLine("      Run QA tests on parts specified in config file");
            Console.WriteLine("      Default config: C:\\Temp\\nm_qa_config.json");
            Console.WriteLine();
            Console.WriteLine("  NM.BatchRunner.exe --pipeline --file <path>");
            Console.WriteLine("      Run processing pipeline on a single file");
            Console.WriteLine();
            Console.WriteLine("  NM.BatchRunner.exe --help");
            Console.WriteLine("      Show this help message");
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
    }
}
