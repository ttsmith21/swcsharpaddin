using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NM.Core;
using NM.Core.DataModel;
using NM.Core.Export;
using NM.SwAddin.Import;
using NM.SwAddin.Pipeline;
using NM.SwAddin.UI;
using NM.SwAddin.Utils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Processing
{
    /// <summary>
    /// Recursively enumerates folder contents, imports neutral files, opens SOLIDWORKS docs safely,
    /// and reports progress. Business logic (validation/processing) is left for orchestrators (vNext).
    /// </summary>
    public sealed class FolderProcessor
    {
        private readonly ISldWorks _swApp;
        private readonly SolidWorksFileOperations _fileOps;
        private readonly StepImportHandler _importer;

        private static readonly string[] PartExts = new[] { ".sldprt", ".sldasm", ".slddrw" };
        private static readonly string[] NeutralExts = new[] { ".step", ".stp", ".iges", ".igs", ".x_t", ".xt" };

        public FolderProcessor(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
            _fileOps = new SolidWorksFileOperations(_swApp);
            _importer = new StepImportHandler(_swApp);
        }

        public sealed class FolderRunResult
        {
            public int TotalDiscovered { get; set; }
            public int ImportedCount { get; set; }
            public int OpenedOk { get; set; }
            public int FailedOpen { get; set; }
            public int Skipped { get; set; }
            public int Processed { get; set; }
            public List<string> Errors { get; } = new List<string>();
            // New: aggregate minimal PartData DTOs for export
            public List<PartData> Parts { get; } = new List<PartData>();
        }

        /// <summary>
        /// Process a folder with a modeless progress dialog. Returns a summary result.
        /// </summary>
        public FolderRunResult ProcessFolder(string root, bool recursive = true)
        {
            const string proc = nameof(ProcessFolder);
            ErrorHandler.PushCallStack(proc);
            var result = new FolderRunResult();
            ProgressForm progress = null;
            try
            {
                if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
                {
                    ErrorHandler.HandleError(proc, $"Invalid folder: {root}");
                    return result;
                }

                var files = EnumerateTargetFiles(root, recursive).ToList();
                result.TotalDiscovered = files.Count;

                progress = new ProgressForm();
                progress.SetMax(files.Count);
                progress.SetSummary($"Folder: {root}");
                progress.Show();
                Application.DoEvents();

                int step = 0;
                bool debugMode = NM.Core.Configuration.Logging.EnableDebugMode;

                // Suppress document display windows during batch opens (avoids SceneGraph overhead).
                // DocumentVisible(false) prevents SolidWorks from allocating map textures for models
                // we're only processing programmatically. Skipped in debug mode so the user can see files open.
                if (!debugMode)
                {
                    try
                    {
                        _swApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocPART);
                        _swApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocASSEMBLY);
                        ErrorHandler.LogInfo("[PERF] FolderProcessor: DocumentVisible=false for batch opens");
                    }
                    catch (Exception dvEx)
                    {
                        ErrorHandler.LogInfo($"[PERF] FolderProcessor: DocumentVisible not available: {dvEx.Message}");
                    }
                }

                // Batch performance optimization: disable graphics updates during file processing loop
                using (new BatchPerformanceScope(_swApp, null))
                using (var swProgress = new SwProgressBar(_swApp, files.Count, "NM Folder Processing"))
                foreach (var file in files)
                {
                    if (progress.IsCanceled || swProgress.UserCanceled) break;
                    step++;
                    var fileName = Path.GetFileName(file);
                    progress.SetStep(step, fileName);
                    swProgress.Update(step, $"Processing {step}/{files.Count}: {fileName}");
                    Application.DoEvents();

                    var ext = (Path.GetExtension(file) ?? string.Empty).ToLowerInvariant();

                    if (NeutralExts.Contains(ext))
                    {
                        string newPath = null;
                        // Heuristic: STEP/IGES/XT could be part or assembly. Try assembly first, fallback to part.
                        newPath = _importer.ImportToAssembly(file);
                        if (string.IsNullOrEmpty(newPath))
                            newPath = _importer.ImportToPart(file);

                        if (!string.IsNullOrEmpty(newPath) && File.Exists(newPath))
                        {
                            result.ImportedCount++;
                            // Continue processing on the new path
                            HandleModelPath(newPath, ref result);
                        }
                        else
                        {
                            result.Errors.Add($"Import failed: {file}");
                            result.FailedOpen++;
                        }
                    }
                    else if (PartExts.Contains(ext))
                    {
                        HandleModelPath(file, ref result);
                    }
                    else
                    {
                        result.Skipped++;
                    }
                }

                // Export aggregated results (CSV + ERP TSV) to the root folder as an initial DTO integration
                try
                {
                    var exporter = new ExportManager();
                    var csvPath = Path.Combine(root, "QuoteSummary.csv");
                    exporter.ExportToCsv(result.Parts, csvPath);

                    var erpPath = Path.Combine(root, "ErpImport.txt");
                    exporter.ExportToErp(result.Parts, erpPath, '\t');
                }
                catch (Exception ex)
                {
                    // Non-fatal: log and continue
                    ErrorHandler.HandleError(proc, "Export failed", ex, ErrorHandler.LogLevel.Warning, root);
                }

                progress.SetSummary($"Done. Found={result.TotalDiscovered}, Imported={result.ImportedCount}, Opened={result.OpenedOk}, Failed={result.FailedOpen}, Skipped={result.Skipped}");
                Application.DoEvents();
                return result;
            }
            finally
            {
                // Restore DocumentVisible for both document types
                try
                {
                    _swApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART);
                    _swApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocASSEMBLY);
                }
                catch { }

                try { if (progress != null && !progress.IsDisposed) progress.Close(); } catch { }
                ErrorHandler.PopCallStack();
            }
        }

        private void HandleModelPath(string path, ref FolderRunResult res)
        {
            const string proc = nameof(HandleModelPath);
            try
            {
                var model = _fileOps.OpenSWDocument(path, silent: true, readOnly: false);
                if (model == null)
                {
                    res.FailedOpen++;
                    return;
                }

                res.OpenedOk++;

                // Route to processing pipeline (single-part or assembly via AutoWorkflow)
                var type = (swDocumentTypes_e)model.GetType();
                if (type == swDocumentTypes_e.swDocPART)
                {
                    var pd = MainRunner.RunSinglePartData(_swApp, model, options: null);
                    res.Parts.Add(pd);

                    if (pd.Status == ProcessingStatus.Success) res.Processed++;
                    else if (pd.Status == ProcessingStatus.Failed) res.Errors.Add(pd.FailureReason ?? "Processing failed");
                }
                else if (type == swDocumentTypes_e.swDocASSEMBLY)
                {
                    // Use auto workflow to run assembly processing over unique parts
                    NM.SwAddin.Pipeline.AutoWorkflow.Run(_swApp);
                    res.Processed++; // best-effort count; summary dialog provides detail
                }

                _fileOps.CloseSWDocument(model);
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex, ErrorHandler.LogLevel.Warning, path);
                res.FailedOpen++;
            }
        }

        private static string Safe(string s) => s ?? string.Empty;

        private static IEnumerable<string> EnumerateTargetFiles(string root, bool recursive)
        {
            var option = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            IEnumerable<string> SafeEnum(string pattern)
            {
                try { return Directory.EnumerateFiles(root, pattern, option); }
                catch { return Array.Empty<string>(); }
            }

            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in new[] { "*.sldprt", "*.sldasm", "*.slddrw", "*.step", "*.stp", "*.iges", "*.igs", "*.x_t", "*.xt" })
                foreach (var f in SafeEnum(p)) set.Add(f);
            return set.OrderBy(s => s, StringComparer.OrdinalIgnoreCase);
        }
    }
}
