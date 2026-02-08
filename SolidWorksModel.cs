using System;
using System.IO;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin
{
    /// <summary>
    /// SolidWorks API service for ModelInfo data management (document lifecycle + property sync).
    /// </summary>
    public class SolidWorksModel : IDisposable
    {
        public ModelInfo Info { get; private set; }
        public ISldWorks Application { get; private set; }
        public IModelDoc2 Document { get; private set; }
        public bool IsDocumentOpen => Document != null;

        public SolidWorksModel(ModelInfo modelInfo, ISldWorks swApp)
        {
            Info = modelInfo ?? throw new ArgumentNullException(nameof(modelInfo));
            Application = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        public bool OpenDocument()
        {
            ErrorHandler.PushCallStack(nameof(OpenDocument));
            try
            {
                if (IsDocumentOpen) return true;

                if (string.IsNullOrWhiteSpace(Info.FilePath) || !File.Exists(Info.FilePath))
                {
                    // Allow unsaved/virtual docs only via active doc (not supported here)
                    ErrorHandler.HandleError(nameof(OpenDocument), $"File not found: {Info.FilePath}", null, ErrorHandler.LogLevel.Warning);
                    Info.CurrentState = ModelInfo.ModelState.Problem;
                    Info.ProblemDescription = "File not found";
                    return false;
                }

                swDocumentTypes_e type = GetDocumentType();
                int err, warn;
                Document = SwDocumentHelper.OpenDocument(Application, Info.FilePath, type, swOpenDocOptions_e.swOpenDocOptions_Silent, out err, out warn);
                if (Document == null)
                {
                    Info.CurrentState = ModelInfo.ModelState.Problem;
                    Info.ProblemDescription = $"Open failed (err={err})";
                    return false;
                }
                Info.CurrentState = ModelInfo.ModelState.Opened;
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(nameof(OpenDocument), "Exception opening document", ex);
                Info.CurrentState = ModelInfo.ModelState.Problem;
                Info.ProblemDescription = ex.Message;
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        public bool SaveDocument()
        {
            ErrorHandler.PushCallStack(nameof(SaveDocument));
            try
            {
                if (!IsDocumentOpen)
                {
                    if (!OpenDocument()) return false;
                }
                bool ok = SwDocumentHelper.SaveDocument(Document, swSaveAsOptions_e.swSaveAsOptions_Silent);
                if (ok)
                {
                    Info.CurrentState = ModelInfo.ModelState.Processed;
                }
                else
                {
                    Info.CurrentState = ModelInfo.ModelState.Problem;
                    Info.ProblemDescription = "Save failed";
                }
                return ok;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        public void CloseDocument()
        {
            ErrorHandler.PushCallStack(nameof(CloseDocument));
            try
            {
                if (IsDocumentOpen)
                {
                    SwDocumentHelper.CloseDocument(Application, Document);
                    Document = null;
                    Info.CurrentState = ModelInfo.ModelState.Idle;
                }
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Attaches this service to an already-open document without opening by path.
        /// </summary>
        public void Attach(IModelDoc2 doc, string configName = null)
        {
            ErrorHandler.PushCallStack(nameof(Attach));
            try
            {
                if (doc == null) throw new ArgumentNullException(nameof(doc));
                Document = doc;
                string path = doc.GetPathName() ?? string.Empty;
                string cfg = configName;
                try
                {
                    cfg = cfg ?? doc.ConfigurationManager?.ActiveConfiguration?.Name ?? string.Empty;
                }
                catch { cfg = cfg ?? string.Empty; }

                // Initialize Info (tolerates empty path)
                Info.Initialize(path, cfg);
                Info.CurrentState = ModelInfo.ModelState.Opened;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        public bool LoadPropertiesFromSolidWorks()
        {
            ErrorHandler.PushCallStack(nameof(LoadPropertiesFromSolidWorks));
            bool openedHere = false;
            try
            {
                if (!IsDocumentOpen)
                {
                    if (!OpenDocument()) return false;
                    openedHere = true;
                }

                string config = Info.ConfigurationName ?? string.Empty;
                ErrorHandler.DebugLog($"Loading properties from config '{config}' (empty = document-level) for '{Document?.GetTitle()}'");
                if (!SwPropertyHelper.GetCustomProperties(Document, config, out var names, out var types, out var values))
                {
                    Info.CurrentState = ModelInfo.ModelState.Problem;
                    Info.ProblemDescription = "Failed to retrieve properties";
                    return false;
                }

                // Reset and load
                Info.CustomProperties.InitializeWithDefaults();
                for (int i = 0; i < names.Length; i++)
                {
                    var name = names[i];
                    var val = (i < values.Length) ? values[i] : string.Empty;
                    var tp = (CustomPropertyType)((i < types.Length) ? types[i] : (int)swCustomInfoType_e.swCustomInfoText);
                    Info.CustomProperties.SetPropertyValue(name, val, tp);
                }
                Info.CustomProperties.MarkClean();
                Info.CurrentState = ModelInfo.ModelState.Validated;
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(nameof(LoadPropertiesFromSolidWorks), "Exception reading properties", ex);
                Info.CurrentState = ModelInfo.ModelState.Problem;
                Info.ProblemDescription = ex.Message;
                return false;
            }
            finally
            {
                if (openedHere) CloseDocument();
                ErrorHandler.PopCallStack();
            }
        }

        public bool SavePropertiesToSolidWorks()
        {
            ErrorHandler.PushCallStack(nameof(SavePropertiesToSolidWorks));
            bool openedHere = false;
            try
            {
                if (!IsDocumentOpen)
                {
                    if (!OpenDocument()) return false;
                    openedHere = true;
                }

                string config = Info.ConfigurationName ?? string.Empty;
                var states = Info.CustomProperties.GetPropertyStates();
                var values = Info.CustomProperties.GetProperties();
                var types = Info.CustomProperties.GetPropertyTypes();
                ErrorHandler.DebugLog($"Saving {states.Count} properties to config '{config}' (empty = document-level) for '{Document?.GetTitle()}'");
                bool ok = true;

                foreach (var kv in states)
                {
                    var name = kv.Key;
                    var state = kv.Value;
                    values.TryGetValue(name, out var objValue);
                    string val = objValue?.ToString() ?? string.Empty;
                    types.TryGetValue(name, out var tp);

                    switch (state)
                    {
                        case PropertyState.Added:
                            ErrorHandler.DebugLog($"Add prop '{name}'='{val}'");
                            ok &= SwPropertyHelper.AddCustomProperty(Document, name, (swCustomInfoType_e)tp, val, config);
                            break;
                        case PropertyState.Modified:
                            // VBA pattern: Always delete-then-add for reliability
                            ErrorHandler.DebugLog($"Modify prop '{name}'='{val}' (using delete+add)");
                            ok &= SwPropertyHelper.AddCustomProperty(Document, name, (swCustomInfoType_e)tp, val, config);
                            break;
                        case PropertyState.Deleted:
                            ErrorHandler.DebugLog($"Delete prop '{name}'");
                            ok &= SwPropertyHelper.DeleteCustomProperty(Document, name, config);
                            break;
                    }
                }

                if (!ok)
                {
                    Info.CurrentState = ModelInfo.ModelState.Problem;
                    Info.ProblemDescription = "Failed to write custom properties";
                    return false;
                }

                Info.CustomProperties.MarkClean();
                Info.IsDirty = false;

                // Attempt to save only if the document has a valid path (avoid Save As prompts on unsaved docs)
                var path = Document?.GetPathName() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(path))
                {
                    if (!SwDocumentHelper.SaveDocument(Document, swSaveAsOptions_e.swSaveAsOptions_Silent))
                    {
                        Info.CurrentState = ModelInfo.ModelState.Problem;
                        Info.ProblemDescription = "Failed to save document";
                        return false;
                    }
                }

                Info.CurrentState = ModelInfo.ModelState.Processed;
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(nameof(SavePropertiesToSolidWorks), "Exception writing properties", ex);
                Info.CurrentState = ModelInfo.ModelState.Problem;
                Info.ProblemDescription = ex.Message;
                return false;
            }
            finally
            {
                if (openedHere) CloseDocument();
                ErrorHandler.PopCallStack();
            }
        }

        public bool ValidateDocument()
        {
            // Basic validate: ensure doc can open
            return OpenDocument();
        }

        public void RefreshDocument()
        {
            if (IsDocumentOpen)
            {
                SwDocumentHelper.ForceRebuildDoc(Document);
            }
        }

        private swDocumentTypes_e GetDocumentType()
        {
            switch (Info.Type)
            {
                case ModelInfo.ModelType.Part: return swDocumentTypes_e.swDocPART;
                case ModelInfo.ModelType.Assembly: return swDocumentTypes_e.swDocASSEMBLY;
                case ModelInfo.ModelType.Drawing: return swDocumentTypes_e.swDocDRAWING;
                default: return swDocumentTypes_e.swDocNONE;
            }
        }

        public void Dispose()
        {
            try { CloseDocument(); } catch { }
        }
    }
}
