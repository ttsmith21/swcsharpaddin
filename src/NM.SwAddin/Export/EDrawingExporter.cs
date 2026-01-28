using System;
using System.IO;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Export
{
    /// <summary>
    /// Exports SolidWorks documents to eDrawings format (.eprt, .easm, .edrw).
    /// Ported from VBA modExport.bas SaveAsEDrawing() function.
    /// </summary>
    public sealed class EDrawingExporter
    {
        private readonly ISldWorks _swApp;

        public sealed class ExportOptions
        {
            /// <summary>
            /// Customer folder name for organizing exports.
            /// </summary>
            public string CustomerFolder { get; set; }

            /// <summary>
            /// Base output path. If null, uses same folder as source file.
            /// </summary>
            public string OutputBasePath { get; set; }

            /// <summary>
            /// Whether to overwrite existing files.
            /// </summary>
            public bool OverwriteExisting { get; set; } = true;
        }

        public sealed class ExportResult
        {
            public bool Success { get; set; }
            public string OutputPath { get; set; }
            public string Message { get; set; }
            public swDocumentTypes_e DocumentType { get; set; }
        }

        public EDrawingExporter(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Exports the active document to eDrawings format.
        /// </summary>
        public ExportResult ExportActiveDocument(ExportOptions options = null)
        {
            var model = _swApp.ActiveDoc as IModelDoc2;
            if (model == null)
            {
                return new ExportResult { Success = false, Message = "No active document" };
            }
            return Export(model, options);
        }

        /// <summary>
        /// Exports a SolidWorks document to eDrawings format.
        /// Parts export as .eprt, assemblies as .easm, drawings as .edrw.
        /// </summary>
        public ExportResult Export(IModelDoc2 model, ExportOptions options = null)
        {
            const string proc = nameof(Export);
            ErrorHandler.PushCallStack(proc);
            options = options ?? new ExportOptions();

            var result = new ExportResult();

            try
            {
                if (model == null)
                {
                    result.Message = "Model is null";
                    return result;
                }

                var docType = (swDocumentTypes_e)model.GetType();
                result.DocumentType = docType;

                // Get document name without extension
                string docTitle = model.GetTitle() ?? "Untitled";
                string docName = Path.GetFileNameWithoutExtension(docTitle);

                // Determine output extension based on document type
                string extension = GetEDrawingExtension(docType);
                if (string.IsNullOrEmpty(extension))
                {
                    result.Message = $"Unsupported document type: {docType}";
                    return result;
                }

                // Build output path
                string outputPath = BuildOutputPath(model, docName, extension, options);

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Check if file exists and overwrite is disabled
                if (File.Exists(outputPath) && !options.OverwriteExisting)
                {
                    result.Success = true;
                    result.OutputPath = outputPath;
                    result.Message = "File already exists (overwrite disabled)";
                    return result;
                }

                // Export to eDrawings format
                int errors = 0;
                int warnings = 0;
                bool saved = model.SaveAs4(
                    outputPath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    ref errors,
                    ref warnings);

                if (saved && errors == 0)
                {
                    result.Success = true;
                    result.OutputPath = outputPath;
                    result.Message = "Export successful";
                }
                else
                {
                    result.Success = false;
                    result.Message = $"Export failed with error code: {errors}";
                }

                return result;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                result.Message = "Exception: " + ex.Message;
                return result;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Exports a document from a file path to eDrawings format.
        /// Opens the document, exports it, then closes it.
        /// </summary>
        public ExportResult ExportFromFile(string filePath, ExportOptions options = null)
        {
            const string proc = nameof(ExportFromFile);
            ErrorHandler.PushCallStack(proc);

            try
            {
                if (!File.Exists(filePath))
                {
                    return new ExportResult { Success = false, Message = "File not found: " + filePath };
                }

                // Determine document type from extension
                var docType = GetDocumentType(filePath);
                if (docType == swDocumentTypes_e.swDocNONE)
                {
                    return new ExportResult { Success = false, Message = "Unknown file type" };
                }

                int errors = 0;
                int warnings = 0;
                var model = _swApp.OpenDoc6(
                    filePath,
                    (int)docType,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                    "",
                    ref errors,
                    ref warnings) as IModelDoc2;

                if (model == null)
                {
                    return new ExportResult { Success = false, Message = $"Failed to open file (error: {errors})" };
                }

                var result = Export(model, options);

                // Close the document
                _swApp.CloseDoc(model.GetTitle());

                return result;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                return new ExportResult { Success = false, Message = "Exception: " + ex.Message };
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Batch exports all SolidWorks files in a folder to eDrawings format.
        /// </summary>
        public BatchExportResult ExportFolder(string folderPath, ExportOptions options = null, Action<string, int, int> progressCallback = null)
        {
            const string proc = nameof(ExportFolder);
            ErrorHandler.PushCallStack(proc);

            var result = new BatchExportResult { FolderPath = folderPath };

            try
            {
                if (!Directory.Exists(folderPath))
                {
                    result.Message = "Folder not found";
                    return result;
                }

                // Get all SolidWorks files
                var files = Directory.GetFiles(folderPath, "*.sld*", SearchOption.TopDirectoryOnly);
                result.TotalFiles = files.Length;

                for (int i = 0; i < files.Length; i++)
                {
                    string file = files[i];
                    string fileName = Path.GetFileName(file);

                    progressCallback?.Invoke(fileName, i + 1, files.Length);

                    // Skip non-document files (like sldmat, sldlfp, etc.)
                    var docType = GetDocumentType(file);
                    if (docType == swDocumentTypes_e.swDocNONE)
                    {
                        result.SkippedFiles.Add(file);
                        continue;
                    }

                    var exportResult = ExportFromFile(file, options);
                    if (exportResult.Success)
                    {
                        result.SuccessCount++;
                        result.ExportedFiles.Add(exportResult.OutputPath);
                    }
                    else
                    {
                        result.FailedCount++;
                        result.FailedFiles.Add(file);
                    }
                }

                result.Success = result.FailedCount == 0;
                result.Message = $"Exported {result.SuccessCount}/{result.TotalFiles} files";
                return result;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                result.Message = "Exception: " + ex.Message;
                return result;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        public sealed class BatchExportResult
        {
            public bool Success { get; set; }
            public string FolderPath { get; set; }
            public int TotalFiles { get; set; }
            public int SuccessCount { get; set; }
            public int FailedCount { get; set; }
            public string Message { get; set; }
            public System.Collections.Generic.List<string> ExportedFiles { get; } = new System.Collections.Generic.List<string>();
            public System.Collections.Generic.List<string> FailedFiles { get; } = new System.Collections.Generic.List<string>();
            public System.Collections.Generic.List<string> SkippedFiles { get; } = new System.Collections.Generic.List<string>();
        }

        private static string GetEDrawingExtension(swDocumentTypes_e docType)
        {
            switch (docType)
            {
                case swDocumentTypes_e.swDocPART:
                    return ".eprt";
                case swDocumentTypes_e.swDocASSEMBLY:
                    return ".easm";
                case swDocumentTypes_e.swDocDRAWING:
                    return ".edrw";
                default:
                    return null;
            }
        }

        private static swDocumentTypes_e GetDocumentType(string filePath)
        {
            string ext = Path.GetExtension(filePath)?.ToLowerInvariant();
            switch (ext)
            {
                case ".sldprt":
                    return swDocumentTypes_e.swDocPART;
                case ".sldasm":
                    return swDocumentTypes_e.swDocASSEMBLY;
                case ".slddrw":
                    return swDocumentTypes_e.swDocDRAWING;
                default:
                    return swDocumentTypes_e.swDocNONE;
            }
        }

        private static string BuildOutputPath(IModelDoc2 model, string docName, string extension, ExportOptions options)
        {
            string basePath;

            if (!string.IsNullOrEmpty(options.OutputBasePath))
            {
                basePath = options.OutputBasePath;
            }
            else
            {
                string modelPath = model.GetPathName();
                basePath = !string.IsNullOrEmpty(modelPath)
                    ? Path.GetDirectoryName(modelPath)
                    : System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
            }

            // Add customer folder if specified
            if (!string.IsNullOrEmpty(options.CustomerFolder))
            {
                basePath = Path.Combine(basePath, options.CustomerFolder);
            }

            return Path.Combine(basePath, docName + extension);
        }
    }
}
