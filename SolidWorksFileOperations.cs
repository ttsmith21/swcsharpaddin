using System;
using System.IO;
using System.Runtime.InteropServices;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SWConfiguration = NM.Core.Configuration;
using SysEnvironment = System.Environment;

namespace NM.SwAddin
{
    /// <summary>
    /// Service class providing SolidWorks file operations with settings management and error translation.
    /// Follows the "thin glue" approach and composes existing static SolidWorksApiWrapper + ErrorHandler.
    /// </summary>
    public sealed class SolidWorksFileOperations : IDisposable
    {
        private readonly ISldWorks _swApp;
        private bool _disposed;

        private bool _settingsInitialized;
        private SolidWorksSettings _original;

        /// <summary>
        /// Holds app and document-related settings we temporarily change during automation.
        /// </summary>
        private sealed class SolidWorksSettings
        {
            public bool AppVisible { get; set; }
            public bool AppUserControl { get; set; }
            public bool AutomaticRebuild { get; set; }
            public bool DisplayDialogsOnRebuildError { get; set; }
            public bool EnableFreezeBar { get; set; }
        }

        public SolidWorksFileOperations(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Attempts to acquire an ISldWorks instance, preferring an existing process.
        /// Note: In add-ins you should pass the app in the constructor instead of calling this.
        /// </summary>
        public static ISldWorks AcquireApplication(bool allowCreate = true, int startupTimeoutSeconds = 30)
        {
            using (new ErrorScope(nameof(AcquireApplication)))
            {
                try
                {
                    ISldWorks app = null;
                    try
                    {
                        app = (ISldWorks)Marshal.GetActiveObject("SldWorks.Application");
                    }
                    catch
                    {
                        app = null;
                    }

                    if (app != null)
                        return app;

                    if (!allowCreate)
                    {
                        ErrorHandler.HandleError(nameof(AcquireApplication), "SolidWorks instance not found and creation is disabled");
                        return null;
                    }

                    var type = Type.GetTypeFromProgID("SldWorks.Application", throwOnError: false);
                    if (type == null)
                    {
                        ErrorHandler.HandleError(nameof(AcquireApplication), "SolidWorks ProgID not found: SldWorks.Application");
                        return null;
                    }

                    app = (ISldWorks)Activator.CreateInstance(type);
                    if (app == null)
                    {
                        ErrorHandler.HandleError(nameof(AcquireApplication), "Could not create SolidWorks application instance");
                        return null;
                    }

                    var startedAt = DateTime.UtcNow;
                    while (!app.StartupProcessCompleted)
                    {
                        if ((DateTime.UtcNow - startedAt).TotalSeconds > startupTimeoutSeconds)
                        {
                            ErrorHandler.HandleError(nameof(AcquireApplication), "Startup timeout");
                            return null;
                        }
                        System.Windows.Forms.Application.DoEvents();
                    }

                    return app;
                }
                catch (Exception ex)
                {
                    ErrorHandler.HandleError(nameof(AcquireApplication), "Exception acquiring SolidWorks application", ex);
                    return null;
                }
            }
        }

        #region Settings Management
        public bool InitializeAppForMode()
        {
            using (new ErrorScope(nameof(InitializeAppForMode)))
            {
                try
                {
                    if (_swApp == null)
                    {
                        ErrorHandler.HandleError(nameof(InitializeAppForMode), "Invalid SolidWorks application reference");
                        return false;
                    }

                    _original = new SolidWorksSettings
                    {
                        AppVisible = _swApp.Visible,
                        AppUserControl = _swApp.UserControl
                    };

                    bool debugMode = SWConfiguration.Logging.EnableDebugMode;
                    _swApp.Visible = debugMode;
                    _swApp.UserControl = debugMode;

                    _settingsInitialized = true;
                    return true;
                }
                catch (Exception ex)
                {
                    ErrorHandler.HandleError(nameof(InitializeAppForMode), "Exception initializing SolidWorks app", ex);
                    return false;
                }
            }
        }

        public bool SaveOriginalSettings(IModelDoc2 model)
        {
            using (new ErrorScope(nameof(SaveOriginalSettings)))
            {
                try
                {
                    if (_swApp == null)
                    {
                        ErrorHandler.HandleError(nameof(SaveOriginalSettings), "Invalid SolidWorks application reference");
                        return false;
                    }
                    if (model == null)
                    {
                        ErrorHandler.HandleError(nameof(SaveOriginalSettings), "Invalid model reference");
                        return false;
                    }

                    var ext = model.Extension;
                    if (ext == null)
                    {
                        ErrorHandler.HandleError(nameof(SaveOriginalSettings), "ModelDocExtension is null");
                        return false;
                    }

                    if (_original == null)
                        _original = new SolidWorksSettings();

                    _original.AppVisible = _swApp.Visible;
                    _original.AppUserControl = _swApp.UserControl;

                    _settingsInitialized = true;
                    return true;
                }
                catch (Exception ex)
                {
                    ErrorHandler.HandleError(nameof(SaveOriginalSettings), "Exception saving original settings", ex);
                    return false;
                }
            }
        }

        public bool ApplyProductionSettings(IModelDoc2 model)
        {
            using (new ErrorScope(nameof(ApplyProductionSettings)))
            {
                try
                {
                    if (_swApp == null)
                    {
                        ErrorHandler.HandleError(nameof(ApplyProductionSettings), "Invalid SolidWorks application reference");
                        return false;
                    }
                    if (model == null)
                    {
                        ErrorHandler.HandleError(nameof(ApplyProductionSettings), "Invalid model reference");
                        return false;
                    }

                    _swApp.Visible = false;
                    _swApp.UserControl = false;
                    return true;
                }
                catch (Exception ex)
                {
                    ErrorHandler.HandleError(nameof(ApplyProductionSettings), "Exception applying production settings", ex);
                    return false;
                }
            }
        }

        public bool RestoreOriginalSettings(IModelDoc2 model)
        {
            using (new ErrorScope(nameof(RestoreOriginalSettings)))
            {
                try
                {
                    if (_swApp == null)
                    {
                        ErrorHandler.HandleError(nameof(RestoreOriginalSettings), "Invalid SolidWorks application reference");
                        return false;
                    }
                    if (model == null)
                    {
                        ErrorHandler.HandleError(nameof(RestoreOriginalSettings), "Invalid model reference");
                        return false;
                    }
                    if (!_settingsInitialized || _original == null)
                    {
                        ErrorHandler.HandleError(nameof(RestoreOriginalSettings), "Settings not initialized", null, ErrorHandler.LogLevel.Warning);
                        return false;
                    }

                    _swApp.Visible = _original.AppVisible;
                    _swApp.UserControl = _original.AppUserControl;

                    _settingsInitialized = false;
                    return true;
                }
                catch (Exception ex)
                {
                    ErrorHandler.HandleError(nameof(RestoreOriginalSettings), "Exception restoring original settings", ex);
                    return false;
                }
            }
        }
        #endregion

        #region File Operations
        public IModelDoc2 OpenSWDocument(string filePath, bool silent = true, bool readOnly = false, string configurationName = "")
        {
            using (new ErrorScope(nameof(OpenSWDocument)))
            {
                bool changedVisibility = false;
                try
                {
                    if (_swApp == null)
                    {
                        ErrorHandler.HandleError(nameof(OpenSWDocument), "Invalid SolidWorks application reference");
                        return null;
                    }
                    if (string.IsNullOrWhiteSpace(filePath))
                    {
                        ErrorHandler.HandleError(nameof(OpenSWDocument), "Invalid file path");
                        return null;
                    }
                    if (!File.Exists(filePath))
                    {
                        ErrorHandler.HandleError(nameof(OpenSWDocument), $"File not found: {filePath}");
                        return null;
                    }

                    var openDoc = GetOpenDocumentByPath(filePath);
                    if (openDoc != null) return openDoc;

                    var docType = DetermineDocType(filePath);
                    var options = swOpenDocOptions_e.swOpenDocOptions_Silent;
                    if (!silent) options = (swOpenDocOptions_e)0;
                    if (readOnly) options |= swOpenDocOptions_e.swOpenDocOptions_ReadOnly;

                    bool origVisible = _swApp.Visible;
                    if (silent && origVisible)
                    {
                        _swApp.Visible = false;
                        changedVisibility = true;
                    }

                    int err, warn;
                    var model = SwDocumentHelper.OpenDocument(_swApp, filePath, docType, options, out err, out warn);
                    if (model == null)
                    {
                        ErrorHandler.HandleError(nameof(OpenSWDocument), $"Failed to open: {filePath}\n{TranslateOpenDocErrors(err, warn)}");
                        return null;
                    }
                    if (err != 0 || warn != 0)
                    {
                        ErrorHandler.HandleError(nameof(OpenSWDocument), $"Opened with issues: {filePath}\n{TranslateOpenDocErrors(err, warn)}", null, ErrorHandler.LogLevel.Warning);
                    }

                    SaveOriginalSettings(model);
                    if (silent) ApplyProductionSettings(model);
                    return model;
                }
                catch (Exception ex)
                {
                    ErrorHandler.HandleError(nameof(OpenSWDocument), $"Exception opening: {filePath}", ex);
                    return null;
                }
                finally
                {
                    if (changedVisibility)
                    {
                        try { _swApp.Visible = true; } catch { }
                    }
                }
            }
        }

        public IModelDoc2 GetOpenDocumentByPath(string filePath)
        {
            using (new ErrorScope(nameof(GetOpenDocumentByPath)))
            {
                try
                {
                    if (_swApp == null) return null;
                    if (string.IsNullOrWhiteSpace(filePath)) return null;
                    string target = filePath.ToLowerInvariant();
                    int count = _swApp.GetDocumentCount();
                    if (count <= 0) return null;
                    var docs = _swApp.GetDocuments() as object[];
                    if (docs == null) return null;
                    for (int i = 0; i < docs.Length; i++)
                    {
                        if (docs[i] is IModelDoc2 m)
                        {
                            string p = (m.GetPathName() ?? string.Empty).ToLowerInvariant();
                            if (string.Equals(p, target, StringComparison.Ordinal))
                                return m;
                        }
                    }
                    return null;
                }
                catch (Exception ex)
                {
                    ErrorHandler.HandleError(nameof(GetOpenDocumentByPath), "Exception enumerating open documents", ex);
                    return null;
                }
            }
        }

        public bool SaveSWDocument(IModelDoc2 model, swSaveAsOptions_e options = swSaveAsOptions_e.swSaveAsOptions_Silent)
        {
            using (new ErrorScope(nameof(SaveSWDocument)))
            {
                try
                {
                    if (!SolidWorksApiWrapper.ValidateModel(model, nameof(SaveSWDocument))) return false;
                    int err = 0, warn = 0;
                    var ok = model.Save3((int)options, ref err, ref warn);
                    if (!ok)
                        ErrorHandler.HandleError(nameof(SaveSWDocument), $"Failed to save: {model.GetTitle()}\n{TranslateSaveDocErrors(err)}");
                    return ok;
                }
                catch (Exception ex)
                {
                    ErrorHandler.HandleError(nameof(SaveSWDocument), "Exception during save", ex);
                    return false;
                }
            }
        }

        public bool CloseSWDocument(IModelDoc2 model)
        {
            using (new ErrorScope(nameof(CloseSWDocument)))
            {
                try
                {
                    if (_swApp == null)
                    {
                        ErrorHandler.HandleError(nameof(CloseSWDocument), "Invalid SolidWorks application reference");
                        return false;
                    }
                    if (!SolidWorksApiWrapper.ValidateModel(model, nameof(CloseSWDocument))) return false;

                    string title = model.GetTitle();
                    _swApp.CloseDoc(title);
                    return true;
                }
                catch (Exception ex)
                {
                    ErrorHandler.HandleError(nameof(CloseSWDocument), "Exception closing document", ex);
                    return false;
                }
            }
        }

        public bool SaveAndCloseSWDocument(IModelDoc2 model, bool silent = true, bool forceClose = true)
        {
            using (new ErrorScope(nameof(SaveAndCloseSWDocument)))
            {
                try
                {
                    if (_swApp == null)
                    {
                        ErrorHandler.HandleError(nameof(SaveAndCloseSWDocument), "Invalid SolidWorks application reference");
                        return false;
                    }
                    if (!SolidWorksApiWrapper.ValidateModel(model, nameof(SaveAndCloseSWDocument))) return false;

                    string path = model.GetPathName();
                    SaveOriginalSettings(model);

                    if (!SaveSWDocument(model, silent ? swSaveAsOptions_e.swSaveAsOptions_Silent : (swSaveAsOptions_e)0))
                    {
                        ErrorHandler.HandleError(nameof(SaveAndCloseSWDocument), $"Failed to save: {path}");
                        if (!forceClose) return false;
                    }

                    if (!CloseSWDocument(model))
                    {
                        ErrorHandler.HandleError(nameof(SaveAndCloseSWDocument), $"Failed to close: {path}");
                        return false;
                    }

                    return true;
                }
                finally
                {
                    try { if (model != null) RestoreOriginalSettings(model); } catch { }
                }
            }
        }

        public bool SaveAs(IModelDoc2 model, string newPath,
            swSaveAsOptions_e options = swSaveAsOptions_e.swSaveAsOptions_Silent,
            swSaveAsVersion_e version = swSaveAsVersion_e.swSaveAsCurrentVersion)
        {
            using (new ErrorScope(nameof(SaveAs)))
            {
                try
                {
                    if (!SolidWorksApiWrapper.ValidateModel(model, nameof(SaveAs))) return false;
                    if (string.IsNullOrWhiteSpace(newPath))
                    {
                        ErrorHandler.HandleError(nameof(SaveAs), "Empty SaveAs path");
                        return false;
                    }

                    var ext = model.Extension;
                    if (ext == null)
                    {
                        ErrorHandler.HandleError(nameof(SaveAs), "ModelDocExtension is null");
                        return false;
                    }

                    var type = ext.GetType();
                    System.Reflection.MethodInfo best = null;
                    foreach (var m in type.GetMethods())
                    {
                        if (!m.Name.StartsWith("SaveAs", StringComparison.OrdinalIgnoreCase)) continue;
                        var ps = m.GetParameters();
                        if (ps.Length == 8 || ps.Length == 6) { best = m; break; }
                        if (ps.Length == 5 || ps.Length == 4) { best = m; break; }
                    }

                    if (best == null)
                    {
                        ErrorHandler.HandleError(nameof(SaveAs), "No SaveAs method found on ModelDocExtension");
                        return false;
                    }

                    var parms = best.GetParameters();
                    var args = new object[parms.Length];
                    for (int i = 0; i < parms.Length; i++)
                    {
                        var p = parms[i];
                        var pt = p.ParameterType.IsByRef ? p.ParameterType.GetElementType() : p.ParameterType;
                        string name = p.Name?.ToLowerInvariant() ?? string.Empty;

                        if (pt == typeof(string))
                            args[i] = newPath;
                        else if (pt == typeof(int) && (name.Contains("version") || i == 1))
                            args[i] = (int)version;
                        else if (pt == typeof(int) && (name.Contains("option") || name.Contains("options")))
                            args[i] = (int)options;
                        else if (pt == typeof(bool))
                            args[i] = true;
                        else if (pt == typeof(object))
                            args[i] = null;
                        else if (p.ParameterType.IsByRef && pt == typeof(int))
                            args[i] = 0;
                        else
                            args[i] = pt.IsValueType ? Activator.CreateInstance(pt) : null;
                    }

                    var result = best.Invoke(ext, args);
                    bool ok = Convert.ToBoolean(result ?? false);
                    if (!ok)
                    {
                        for (int i = 0; i < parms.Length; i++)
                        {
                            var p = parms[i];
                            var pt = p.ParameterType.IsByRef ? p.ParameterType.GetElementType() : p.ParameterType;
                            if (pt == typeof(int) && (p.Name?.ToLowerInvariant().Contains("version") == true)) args[i] = (int)version;
                            else if (pt == typeof(int) && (p.Name?.ToLowerInvariant().Contains("option") == true)) args[i] = (int)options;
                        }
                        result = best.Invoke(ext, args);
                        ok = Convert.ToBoolean(result ?? false);
                    }

                    if (!ok)
                    {
                        ErrorHandler.HandleError(nameof(SaveAs), $"SaveAs failed via '{best.Name}' for '{newPath}'");
                        return false;
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    ErrorHandler.HandleError(nameof(SaveAs), $"Exception SaveAs: {newPath}", ex);
                    return false;
                }
            }
        }
        #endregion

        #region Helpers
        private static swDocumentTypes_e DetermineDocType(string filePath)
        {
            switch ((Path.GetExtension(filePath) ?? string.Empty).ToLowerInvariant())
            {
                case ".sldprt": return swDocumentTypes_e.swDocPART;
                case ".sldasm": return swDocumentTypes_e.swDocASSEMBLY;
                case ".slddrw": return swDocumentTypes_e.swDocDRAWING;
                default: return swDocumentTypes_e.swDocNONE;
            }
        }

        private static string TranslateOpenDocErrors(int errors, int warnings)
        {
            var msg = string.Empty;
            if (errors != 0)
            {
                if ((errors & (int)swFileLoadError_e.swGenericError) != 0) msg += "General error occurred during file open." + SysEnvironment.NewLine;
                if ((errors & (int)swFileLoadError_e.swFileNotFoundError) != 0) msg += "File not found." + SysEnvironment.NewLine;
                if ((errors & (int)swFileLoadError_e.swFileRequiresRepairError) != 0) msg += "File requires repair." + SysEnvironment.NewLine;
                // No generic 'already open' error flag present in some versions; rely on warning instead.
            }
            if (warnings != 0)
            {
                if ((warnings & (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen) != 0) msg += "Warning: File is already open." + SysEnvironment.NewLine;
            }
            return msg;
        }

        private static string TranslateSaveDocErrors(int errors)
        {
            var msg = string.Empty;
            if (errors != 0)
            {
                if ((errors & (int)swFileSaveError_e.swGenericSaveError) != 0) msg += "General error occurred while saving." + SysEnvironment.NewLine;
                // Some save error flags vary by version; only include the generic one for portability.
            }
            return msg;
        }
        #endregion

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            // Nothing to release here because _swApp is owned by SolidWorks; do not ReleaseComObject.
            // Settings are restored per-operation; no global state to unwind.
        }
    }
}
