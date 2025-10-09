using System;
using System.IO;
using NM.Core;
using NM.Core.Models;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using NM.SwAddin; // for SolidWorksFileOperations

namespace NM.SwAddin.Services
{
    /// <summary>
    /// Thin wrapper for opening/saving/closing docs. Main-thread only.
    /// </summary>
    public sealed class DocumentService
    {
        private readonly ISldWorks _swApp;
        public DocumentService(ISldWorks swApp) { _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp)); }

        public SwModelInfo GetActiveModel()
        {
            var model = _swApp?.IActiveDoc2; if (model == null) return null;

            // Try to get a real path; unsaved docs return empty. We still want a model wrapper.
            var path = string.Empty;
            try { path = model.GetPathName(); } catch { path = string.Empty; }

            if (string.IsNullOrWhiteSpace(path))
            {
                // Build a synthetic path with the correct extension so SwModelInfo can infer type.
                string title = string.Empty; try { title = model.GetTitle() ?? string.Empty; } catch { }
                int docType = -1; try { docType = model.GetType(); } catch { docType = -1; }
                string ext = ".sldprt";
                if (docType == (int)swDocumentTypes_e.swDocASSEMBLY) ext = ".sldasm";
                else if (docType == (int)swDocumentTypes_e.swDocDRAWING) ext = ".slddrw";

                if (!string.IsNullOrWhiteSpace(title))
                {
                    var tl = title.ToLowerInvariant();
                    if (tl.EndsWith(".sldprt") || tl.EndsWith(".sldasm") || tl.EndsWith(".slddrw"))
                        path = title;
                    else
                        path = title + ext;
                }
                else
                {
                    path = "Unsaved" + ext;
                }
            }

            var config = string.Empty;
            try { config = model.IGetActiveConfiguration()?.Name ?? string.Empty; } catch { config = string.Empty; }

            var info = new SwModelInfo(path, config) { ModelDoc = model };
            return info;
        }

        public bool Open(SwModelInfo info, bool silent = true)
        {
            if (info == null) throw new ArgumentNullException(nameof(info));
            if (info.ModelDoc != null) return true;
            if (!File.Exists(info.FilePath)) { ErrorHandler.HandleError(nameof(DocumentService), $"File not found: {info.FilePath}", null); return false; }
            int errors = 0, warnings = 0; int type = (int)swDocumentTypes_e.swDocNONE;
            if (info.Type == SwModelInfo.ModelType.Part) type = (int)swDocumentTypes_e.swDocPART;
            else if (info.Type == SwModelInfo.ModelType.Assembly) type = (int)swDocumentTypes_e.swDocASSEMBLY;
            else if (info.Type == SwModelInfo.ModelType.Drawing) type = (int)swDocumentTypes_e.swDocDRAWING;

            PerformanceTracker.Instance.StartTimer("OpenDocument");
            try
            {
                var opts = silent ? (int)swOpenDocOptions_e.swOpenDocOptions_Silent : (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel;
                var model = _swApp.OpenDoc6(info.FilePath, type, opts, info.Configuration, ref errors, ref warnings) as IModelDoc2;
                if (model == null) { ErrorHandler.HandleError(nameof(DocumentService), $"Failed to open {info.FilePath}. Errors={errors} Warnings={warnings}", null); return false; }
                info.ModelDoc = model; if (!string.IsNullOrWhiteSpace(info.Configuration)) model.ShowConfiguration2(info.Configuration);
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(nameof(DocumentService), $"Exception opening {info.FileName}", ex);
                return false;
            }
            finally { PerformanceTracker.Instance.StopTimer("OpenDocument"); }
        }

        public bool Save(SwModelInfo info)
        {
            if (info?.ModelDoc == null) return false; var model = info.ModelDoc as IModelDoc2; if (model == null) return false;
            PerformanceTracker.Instance.StartTimer("SaveDocument");
            try
            {
                string path = string.Empty; try { path = model.GetPathName(); } catch { path = string.Empty; }
                if (string.IsNullOrWhiteSpace(path))
                {
                    var ops = new SolidWorksFileOperations(_swApp);
                    var okTemp = ops.SaveSWDocument(model, swSaveAsOptions_e.swSaveAsOptions_Silent);
                    if (!okTemp)
                    {
                        ErrorHandler.HandleError(nameof(DocumentService), $"Temp SaveAs failed for {info.FileName}");
                        return false;
                    }
                    // Model saved; do not clear property cache here.
                    info.MarkModelClean();
                    return true;
                }

                int e = 0, w = 0; bool ok = model.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref e, ref w);
                if (!ok) ErrorHandler.HandleError(nameof(DocumentService), $"Save failed for {info.FileName}. Errors={e} Warnings={w}", null);
                if (ok) info.MarkModelClean();
                return ok;
            }
            catch (Exception ex) { ErrorHandler.HandleError(nameof(DocumentService), $"Exception saving {info.FileName}", ex); return false; }
            finally { PerformanceTracker.Instance.StopTimer("SaveDocument"); }
        }

        public bool Close(SwModelInfo info, bool save = false)
        {
            if (info?.ModelDoc == null) return true; var model = info.ModelDoc as IModelDoc2; if (model == null) return false;
            PerformanceTracker.Instance.StartTimer("CloseDocument");
            try
            {
                if (save && info.IsDirty) Save(info);
                var title = model.GetTitle(); _swApp.CloseDoc(title); info.ModelDoc = null; return true;
            }
            catch (Exception ex) { ErrorHandler.HandleError(nameof(DocumentService), $"Exception closing {info.FileName}", ex); return false; }
            finally { PerformanceTracker.Instance.StopTimer("CloseDocument"); }
        }
    }
}
