using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NM.Core;
using NM.Core.DataModel;
using NM.Core.Export;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Pipeline
{
    /// <summary>
    /// Quote workflow orchestration for batch processing imported files.
    /// Ported from VBA SP.bas QuoteStart() and QuoteStartASM() functions.
    /// </summary>
    public sealed class QuoteWorkflow
    {
        private readonly ISldWorks _swApp;
        private readonly string _workingFolder;

        public sealed class QuoteResult
        {
            public bool Success { get; set; }
            public int TotalFiles { get; set; }
            public int ProcessedCount { get; set; }
            public int FailedCount { get; set; }
            public List<string> FailedFiles { get; } = new List<string>();
            public List<string> ProcessedFiles { get; } = new List<string>();
            public List<PartData> ProcessedParts { get; } = new List<PartData>();
            public string Message { get; set; }
            public string ErpExportPath { get; set; }
        }

        public sealed class QuoteOptions
        {
            public bool ProcessAssemblies { get; set; } = false;
            public bool HideSolidWorks { get; set; } = true;
            public bool DeleteOriginalImports { get; set; } = false;
            public Action<string, int, int> ProgressCallback { get; set; }

            // ERP Export options
            public bool GenerateErpExport { get; set; } = false;
            public string ErpExportPath { get; set; }
            public string Customer { get; set; } = "";
            public string ParentPartNumber { get; set; } = "";

            // Processing options to pass to MainRunner
            public ProcessingOptions ProcessingOptions { get; set; }
        }

        private static readonly string[] ImportExtensions = { ".igs", ".iges", ".step", ".stp", ".sat" };

        public QuoteWorkflow(ISldWorks swApp, string workingFolder)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
            _workingFolder = workingFolder ?? throw new ArgumentNullException(nameof(workingFolder));
        }

        /// <summary>
        /// Runs quote workflow for parts only (QuoteStart).
        /// Imports STEP/IGES/SAT files, converts to SLDPRT, then processes each part.
        /// </summary>
        public QuoteResult RunPartsQuote(QuoteOptions options = null)
        {
            const string proc = nameof(RunPartsQuote);
            ErrorHandler.PushCallStack(proc);
            options = options ?? new QuoteOptions();

            var result = new QuoteResult();

            try
            {
                // Phase 1: Import foreign files and convert to SLDPRT/SLDASM
                var importedFiles = ImportForeignFiles(options);

                // Phase 2: Collect all SLDPRT files for processing
                var partFiles = Directory.GetFiles(_workingFolder, "*.sldprt", SearchOption.TopDirectoryOnly)
                    .ToList();

                result.TotalFiles = partFiles.Count;

                if (partFiles.Count == 0)
                {
                    result.Success = true;
                    result.Message = "No part files found to process";
                    return result;
                }

                // Phase 3: Process each part
                if (options.HideSolidWorks)
                {
                    _swApp.Visible = false;
                    _swApp.UserControl = false;
                }

                var procOptions = options.ProcessingOptions ?? new ProcessingOptions();

                for (int i = 0; i < partFiles.Count; i++)
                {
                    string partFile = partFiles[i];
                    string partName = Path.GetFileNameWithoutExtension(partFile);

                    options.ProgressCallback?.Invoke(partName, i + 1, partFiles.Count);

                    var partData = ProcessSinglePartData(partFile, procOptions);
                    if (partData != null && partData.Status == ProcessingStatus.Success)
                    {
                        result.ProcessedCount++;
                        result.ProcessedFiles.Add(partFile);
                        result.ProcessedParts.Add(partData);
                    }
                    else
                    {
                        result.FailedCount++;
                        result.FailedFiles.Add(partFile);
                    }
                }

                // Phase 4: Generate ERP export if requested
                if (options.GenerateErpExport && result.ProcessedParts.Count > 0)
                {
                    try
                    {
                        string exportPath = options.ErpExportPath;
                        if (string.IsNullOrEmpty(exportPath))
                        {
                            exportPath = Path.Combine(_workingFolder, "Import.prn");
                        }

                        string parentNumber = options.ParentPartNumber;
                        if (string.IsNullOrEmpty(parentNumber))
                        {
                            parentNumber = Path.GetFileName(_workingFolder);
                        }

                        var erpData = ErpExportDataBuilder.FromPartDataCollection(
                            result.ProcessedParts,
                            parentNumber,
                            options.Customer,
                            $"Quote for {parentNumber}");

                        var exporter = new ErpExportFormat { Customer = options.Customer };
                        exporter.ExportToImportPrn(erpData, exportPath);
                        result.ErpExportPath = exportPath;
                    }
                    catch (Exception ex)
                    {
                        ErrorHandler.HandleError(proc, "ERP export failed", ex);
                    }
                }

                result.Success = result.FailedCount == 0;
                result.Message = $"Processed {result.ProcessedCount}/{result.TotalFiles} parts";
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                result.Success = false;
                result.Message = "Quote workflow exception: " + ex.Message;
            }
            finally
            {
                // Restore SolidWorks visibility
                _swApp.Visible = true;
                _swApp.UserControl = true;
                ErrorHandler.PopCallStack();
            }

            return result;
        }

        /// <summary>
        /// Runs quote workflow for assemblies (QuoteStartASM).
        /// Imports files, then processes each as an assembly with BOM extraction.
        /// </summary>
        public QuoteResult RunAssemblyQuote(QuoteOptions options = null)
        {
            const string proc = nameof(RunAssemblyQuote);
            ErrorHandler.PushCallStack(proc);
            options = options ?? new QuoteOptions { ProcessAssemblies = true };

            var result = new QuoteResult();

            try
            {
                // Phase 1: Import foreign files
                var importedFiles = ImportForeignFiles(options);

                result.TotalFiles = importedFiles.Count;

                if (importedFiles.Count == 0)
                {
                    result.Success = true;
                    result.Message = "No files imported";
                    return result;
                }

                // Phase 2: Process each imported file
                if (options.HideSolidWorks)
                {
                    _swApp.Visible = false;
                    _swApp.UserControl = false;
                }

                for (int i = 0; i < importedFiles.Count; i++)
                {
                    var imported = importedFiles[i];
                    options.ProgressCallback?.Invoke(imported.FileName, i + 1, importedFiles.Count);

                    bool processed = false;
                    if (imported.DocType == swDocumentTypes_e.swDocASSEMBLY)
                    {
                        processed = ProcessAssembly(imported.FilePath);
                    }
                    else
                    {
                        processed = ProcessSinglePart(imported.FilePath);
                    }

                    if (processed)
                    {
                        result.ProcessedCount++;
                        result.ProcessedFiles.Add(imported.FilePath);
                    }
                    else
                    {
                        result.FailedCount++;
                        result.FailedFiles.Add(imported.FilePath);
                    }
                }

                result.Success = result.FailedCount == 0;
                result.Message = $"Processed {result.ProcessedCount}/{result.TotalFiles} files";
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                result.Success = false;
                result.Message = "Assembly quote exception: " + ex.Message;
            }
            finally
            {
                _swApp.Visible = true;
                _swApp.UserControl = true;
                ErrorHandler.PopCallStack();
            }

            return result;
        }

        private sealed class ImportedFile
        {
            public string FilePath { get; set; }
            public string FileName { get; set; }
            public swDocumentTypes_e DocType { get; set; }
        }

        private List<ImportedFile> ImportForeignFiles(QuoteOptions options)
        {
            var imported = new List<ImportedFile>();

            var foreignFiles = Directory.GetFiles(_workingFolder)
                .Where(f => ImportExtensions.Contains(Path.GetExtension(f).ToLowerInvariant()))
                .ToList();

            foreach (var foreignFile in foreignFiles)
            {
                try
                {
                    int errors = 0;
                    int warnings = 0;

                    // Load the foreign file
                    var model = _swApp.LoadFile4(foreignFile, "r", null, ref errors) as IModelDoc2;
                    if (model == null) continue;

                    var docType = (swDocumentTypes_e)model.GetType();
                    string docName = model.GetTitle();
                    if (string.IsNullOrEmpty(docName))
                        docName = Path.GetFileNameWithoutExtension(foreignFile);

                    // Remove any extension that might be in the title
                    if (docName.Contains("."))
                        docName = Path.GetFileNameWithoutExtension(docName);

                    string newPath;
                    if (docType == swDocumentTypes_e.swDocASSEMBLY)
                    {
                        newPath = Path.Combine(_workingFolder, docName + ".sldasm");
                    }
                    else
                    {
                        newPath = Path.Combine(_workingFolder, docName + ".sldprt");
                    }

                    // Save as native format
                    int saveErrors = 0;
                    int saveWarnings = 0;
                    bool saved = model.SaveAs4(
                        newPath,
                        (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                        ref saveErrors,
                        ref saveWarnings);

                    if (saved)
                    {
                        imported.Add(new ImportedFile
                        {
                            FilePath = newPath,
                            FileName = docName,
                            DocType = docType
                        });
                    }

                    _swApp.CloseAllDocuments(true);
                }
                catch (Exception ex)
                {
                    ErrorHandler.DebugLog($"Failed to import {foreignFile}: {ex.Message}");
                }
            }

            return imported;
        }

        private bool ProcessSinglePart(string partPath)
        {
            var partData = ProcessSinglePartData(partPath, new ProcessingOptions());
            return partData != null && partData.Status == ProcessingStatus.Success;
        }

        private PartData ProcessSinglePartData(string partPath, ProcessingOptions procOptions)
        {
            try
            {
                int errors = 0;
                int warnings = 0;

                var model = _swApp.OpenDoc6(
                    partPath,
                    (int)swDocumentTypes_e.swDocPART,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                    "",
                    ref errors,
                    ref warnings) as IModelDoc2;

                if (model == null)
                {
                    return new PartData
                    {
                        FilePath = partPath,
                        Status = ProcessingStatus.Failed,
                        FailureReason = "Failed to open part file"
                    };
                }

                // Run the main processing pipeline and get PartData
                var partData = MainRunner.RunSinglePartData(_swApp, model, procOptions);

                // Save if successful
                if (partData.Status == ProcessingStatus.Success)
                {
                    model.Save3(
                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                        ref errors,
                        ref warnings);
                }

                _swApp.CloseAllDocuments(true);
                return partData;
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"ProcessSinglePartData failed: {ex.Message}");
                return new PartData
                {
                    FilePath = partPath,
                    Status = ProcessingStatus.Failed,
                    FailureReason = ex.Message
                };
            }
        }

        private bool ProcessAssembly(string assyPath)
        {
            try
            {
                int errors = 0;
                int warnings = 0;

                var model = _swApp.OpenDoc6(
                    assyPath,
                    (int)swDocumentTypes_e.swDocASSEMBLY,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                    "",
                    ref errors,
                    ref warnings) as IModelDoc2;

                if (model == null) return false;

                // For assemblies, we need to process each component
                var assyDoc = model as IAssemblyDoc;
                if (assyDoc == null)
                {
                    _swApp.CloseAllDocuments(true);
                    return false;
                }

                // Get all components and process unique parts
                var config = model.ConfigurationManager?.ActiveConfiguration;
                if (config == null)
                {
                    _swApp.CloseAllDocuments(true);
                    return false;
                }

                var rootComp = config.GetRootComponent3(true) as IComponent2;
                if (rootComp != null)
                {
                    ProcessComponentTree(rootComp);
                }

                // Save assembly
                model.Save3(
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    ref errors,
                    ref warnings);

                _swApp.CloseAllDocuments(true);
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"ProcessAssembly failed: {ex.Message}");
                return false;
            }
        }

        private void ProcessComponentTree(IComponent2 comp)
        {
            if (comp == null) return;

            var childrenObj = comp.GetChildren();
            if (childrenObj == null) return;

            var children = childrenObj as object[];
            if (children == null) return;

            var processedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var childObj in children)
            {
                var child = childObj as IComponent2;
                if (child == null || child.IsSuppressed()) continue;

                string childPath = child.GetPathName();
                if (string.IsNullOrEmpty(childPath)) continue;

                // Only process each unique part once
                if (processedPaths.Contains(childPath)) continue;
                processedPaths.Add(childPath);

                // Get the model doc for this component
                var childModel = child.GetModelDoc2() as IModelDoc2;
                if (childModel != null)
                {
                    var docType = (swDocumentTypes_e)childModel.GetType();
                    if (docType == swDocumentTypes_e.swDocPART)
                    {
                        // Process as part
                        MainRunner.RunSinglePart(_swApp, childModel, new ProcessingOptions());
                    }
                    else if (docType == swDocumentTypes_e.swDocASSEMBLY)
                    {
                        // Recurse into sub-assembly
                        var subConfig = childModel.ConfigurationManager?.ActiveConfiguration;
                        var subRoot = subConfig?.GetRootComponent3(true) as IComponent2;
                        if (subRoot != null)
                        {
                            ProcessComponentTree(subRoot);
                        }
                    }
                }

                // Also recurse into children
                ProcessComponentTree(child);
            }
        }
    }
}
