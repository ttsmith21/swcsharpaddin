using System;
using System.IO;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Import
{
    /// <summary>
    /// Handles importing neutral files (STEP/IGES/XT) into SolidWorks and saving as native .sldprt/.sldasm.
    /// </summary>
    public sealed class StepImportHandler
    {
        private readonly ISldWorks _swApp;

        public StepImportHandler(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Opens a neutral file and saves it as a SolidWorks part next to the source file.
        /// Returns the new .sldprt path or null on failure. For assemblies, prefer ImportToAssembly.
        /// </summary>
        public string ImportToPart(string sourcePath, bool overwrite = false)
        {
            const string proc = nameof(ImportToPart);
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
                {
                    ErrorHandler.HandleError(proc, $"Invalid source path: {sourcePath}");
                    return null;
                }

                int errs = 0, warns = 0;
                IModelDoc2 model = _swApp.OpenDoc6(sourcePath, (int)swDocumentTypes_e.swDocPART,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent, string.Empty, ref errs, ref warns);
                if (errs != 0 || warns != 0)
                    ErrorHandler.DebugLog($"[{proc}] OpenDoc6 returned errs={errs}, warns={warns} (non-fatal for neutral files)");
                if (model == null)
                {
                    ErrorHandler.HandleError(proc, $"Open failed: {sourcePath} (errs={errs}, warns={warns})");
                    return null;
                }

                var fileOps = new SolidWorksFileOperations(_swApp);
                string target = Path.ChangeExtension(sourcePath, ".sldprt");
                if (!overwrite && File.Exists(target))
                {
                    var baseName = Path.GetFileNameWithoutExtension(sourcePath);
                    var dir = Path.GetDirectoryName(sourcePath) ?? string.Empty;
                    target = Path.Combine(dir, baseName + "_import.sldprt");
                }

                bool ok = fileOps.SaveAs(model, target);
                try { fileOps.CloseSWDocument(model); } catch { }

                return ok ? target : null;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                return null;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        /// <summary>
        /// Opens a neutral assembly (STEP/IGES/XT). Saves as a native SolidWorks assembly with externalized components.
        /// Falls back to part import if assembly import fails.
        /// Returns the new .sldasm or .sldprt path, or null on failure.
        /// </summary>
        public string ImportToAssembly(string sourcePath, bool overwrite = false)
        {
            const string proc = nameof(ImportToAssembly);
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
                {
                    ErrorHandler.HandleError(proc, $"Invalid source path: {sourcePath}");
                    return null;
                }

                int errs = 0, warns = 0;

                // Use LoadFile4 with GetImportFileData — the proper API for neutral file import.
                // This lets SolidWorks decide whether to open as assembly or part.
                Console.WriteLine($"[{proc}] Loading STEP via LoadFile4: {sourcePath}");
                var importData = _swApp.GetImportFileData(sourcePath);
                IModelDoc2 model = _swApp.LoadFile4(sourcePath, "r", importData, ref errs);
                Console.WriteLine($"[{proc}] LoadFile4 result: model={model != null}, errs={errs}");

                // Fallback: check ActiveDoc (some SW versions load successfully but return null)
                if (model == null)
                {
                    model = _swApp.ActiveDoc as IModelDoc2;
                    if (model != null)
                        Console.WriteLine($"[{proc}] Retrieved model from ActiveDoc: {model.GetTitle()}");
                }

                if (model == null)
                {
                    ErrorHandler.HandleError(proc, $"LoadFile4 failed: {sourcePath} (errs={errs})");
                    return null;
                }

                var docType = (swDocumentTypes_e)model.GetType();
                ErrorHandler.DebugLog($"[{proc}] Opened as {docType}: {model.GetTitle()}");

                var fileOps = new SolidWorksFileOperations(_swApp);

                // If SW opened as a part, save as .sldprt
                if (docType == swDocumentTypes_e.swDocPART)
                {
                    string targetPrt = Path.ChangeExtension(sourcePath, ".sldprt");
                    if (!overwrite && File.Exists(targetPrt))
                    {
                        var baseName = Path.GetFileNameWithoutExtension(sourcePath);
                        var dir2 = Path.GetDirectoryName(sourcePath) ?? string.Empty;
                        targetPrt = Path.Combine(dir2, baseName + "_import.sldprt");
                    }
                    bool okPrt = fileOps.SaveAs(model, targetPrt);
                    try { fileOps.CloseSWDocument(model); } catch { }
                    return okPrt ? targetPrt : null;
                }

                // Assembly path
                string targetAsm = Path.ChangeExtension(sourcePath, ".sldasm");
                if (!overwrite && File.Exists(targetAsm))
                {
                    var baseName = Path.GetFileNameWithoutExtension(sourcePath);
                    var dir = Path.GetDirectoryName(sourcePath) ?? string.Empty;
                    targetAsm = Path.Combine(dir, baseName + "_import.sldasm");
                }

                // Save the top-level assembly
                if (!fileOps.SaveAs(model, targetAsm))
                {
                    try { fileOps.CloseSWDocument(model); } catch { }
                    return null;
                }

                // Externalize and save all virtual components under the top-level assembly
                try
                {
                    var asm = model as IAssemblyDoc;
                    if (asm != null)
                    {
                        var compObjs = asm.GetComponents(false) as object[];
                        if (compObjs != null)
                        {
                            ErrorHandler.DebugLog($"[{proc}] Externalizing {compObjs.Length} components...");
                            foreach (var o in compObjs)
                            {
                                var c = o as IComponent2; if (c == null) continue;
                                var child = c.GetModelDoc2() as IModelDoc2; if (child == null) continue;

                                string childPath = child.GetPathName() ?? string.Empty;
                                bool isVirtual = childPath.IndexOf('^') >= 0 || string.IsNullOrWhiteSpace(childPath);
                                if (isVirtual)
                                {
                                    string dir = Path.GetDirectoryName(targetAsm) ?? string.Empty;
                                    string name = (c.Name2 ?? "Component").Replace(":", "_").Replace("*", "_");
                                    string ext = child is IPartDoc ? ".sldprt" : ".sldasm";
                                    string outPath = Path.Combine(dir, name + ext);

                                    try
                                    {
                                        child.Extension.SaveAs(outPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                            (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref errs, ref warns);
                                        ErrorHandler.DebugLog($"[{proc}] Externalized: {name}{ext}");
                                    }
                                    catch { }
                                }
                                else
                                {
                                    ErrorHandler.DebugLog($"[{proc}] Already external: {Path.GetFileName(childPath)}");
                                }
                            }
                        }
                    }
                }
                catch { }

                try { fileOps.CloseSWDocument(model); } catch { }
                return targetAsm;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                return null;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        public static bool IsNeutral(string path)
        {
            try
            {
                var ext = (Path.GetExtension(path) ?? string.Empty).ToLowerInvariant();
                return ext == ".step" || ext == ".stp" || ext == ".iges" || ext == ".igs" || ext == ".x_t" || ext == ".xt";
            }
            catch { return false; }
        }
    }
}
