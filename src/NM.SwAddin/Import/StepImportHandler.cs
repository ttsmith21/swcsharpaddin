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

                if (model == null || errs != 0)
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
        /// Returns the new .sldasm path or null on failure.
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
                // Let SW open the neutral file as an assembly when applicable
                IModelDoc2 model = _swApp.OpenDoc6(sourcePath, (int)swDocumentTypes_e.swDocASSEMBLY,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent, string.Empty, ref errs, ref warns);
                if (model == null || errs != 0)
                {
                    ErrorHandler.HandleError(proc, $"Open failed: {sourcePath} (errs={errs}, warns={warns})");
                    return null;
                }

                var fileOps = new SolidWorksFileOperations(_swApp);
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

                                    // Save child out as an external file
                                    try
                                    {
                                        child.Extension.SaveAs(outPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                            (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref errs, ref warns);
                                    }
                                    catch { }
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
