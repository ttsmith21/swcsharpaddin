using System;
using System.IO;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin
{
    /// <summary>
    /// File operations helper for SolidWorks documents. Thin wrappers that centralize
    /// open/save/close/activate logic and small utilities.
    /// </summary>
    public static class SolidWorksFileOps
    {
        public static IModelDoc2 Open(ISldWorks app, string filePath, swOpenDocOptions_e options, out int err, out int warn)
        {
            const string proc = "SolidWorksFileOps.Open";
            ErrorHandler.PushCallStack(proc);
            try
            {
                err = 0; warn = 0;
                if (app == null) { ErrorHandler.HandleError(proc, "Null ISldWorks"); return null; }
                if (string.IsNullOrWhiteSpace(filePath)) { ErrorHandler.HandleError(proc, "Empty file path"); return null; }

                var docType = DetermineDocTypeByPath(filePath);
                if (docType == swDocumentTypes_e.swDocNONE)
                {
                    ErrorHandler.HandleError(proc, $"Unknown document type for path: {filePath}");
                    return null;
                }

                return SolidWorksApiWrapper.OpenDocument(app, filePath, docType, options, out err, out warn);
            }
            catch (Exception ex)
            {
                err = -1; warn = 0;
                ErrorHandler.HandleError(proc, $"Exception opening: {filePath}", ex);
                return null;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        public static bool EnsureOpen(ISldWorks app, string filePath, out IModelDoc2 doc, bool readOnly = false)
        {
            const string proc = "SolidWorksFileOps.EnsureOpen";
            ErrorHandler.PushCallStack(proc);
            try
            {
                doc = TryGetOpenDoc(app, filePath);
                if (doc != null) return true;

                var opts = swOpenDocOptions_e.swOpenDocOptions_Silent;
                if (readOnly) opts |= swOpenDocOptions_e.swOpenDocOptions_ReadOnly;

                var opened = Open(app, filePath, opts, out var err, out var warn);
                doc = opened;
                if (opened == null)
                {
                    ErrorHandler.HandleError(proc, $"Open failed for {filePath} (err={err}, warn={warn})");
                    return false;
                }
                return true;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        public static IModelDoc2 TryGetOpenDoc(ISldWorks app, string pathOrTitle)
        {
            const string proc = "SolidWorksFileOps.TryGetOpenDoc";
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (app == null || string.IsNullOrWhiteSpace(pathOrTitle)) return null;
                // SolidWorks supports querying by full path or title
                var doc = app.GetOpenDocumentByName(pathOrTitle) as IModelDoc2;
                if (doc != null) return doc;
                var title = Path.GetFileName(pathOrTitle);
                if (!string.IsNullOrEmpty(title))
                {
                    doc = app.GetOpenDocumentByName(title) as IModelDoc2;
                }
                return doc;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, $"Exception resolving open doc: {pathOrTitle}", ex);
                return null;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        public static bool Save(IModelDoc2 doc, swSaveAsOptions_e options = swSaveAsOptions_e.swSaveAsOptions_Silent)
        {
            return SolidWorksApiWrapper.SaveDocument(doc, options);
        }

        public static bool SaveAs(IModelDoc2 doc, string newPath,
            swSaveAsOptions_e options = swSaveAsOptions_e.swSaveAsOptions_Silent,
            swSaveAsVersion_e version = swSaveAsVersion_e.swSaveAsCurrentVersion)
        {
            const string proc = "SolidWorksFileOps.SaveAs";
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(doc, proc)) return false;
                if (string.IsNullOrWhiteSpace(newPath))
                {
                    ErrorHandler.HandleError(proc, "Empty SaveAs path");
                    return false;
                }

                var ext = doc.Extension;
                if (ext == null)
                {
                    ErrorHandler.HandleError(proc, "ModelDocExtension is null");
                    return false;
                }

                int err = 0, warn = 0;
                bool ok = false;
                // Some interop versions expose SaveAs3, others SaveAs4/5. Use reflection to remain compatible.
                var mi = ext.GetType().GetMethod("SaveAs4");
                if (mi != null)
                {
                    var res = mi.Invoke(ext, new object[] { newPath, (int)version, (int)options, null, err, warn });
                    // Some interop map ref params differently; fall back to Save3 if needed
                    ok = Convert.ToBoolean(res ?? false);
                }
                else
                {
                    mi = ext.GetType().GetMethod("SaveAs3");
                    if (mi != null)
                    {
                        var res = mi.Invoke(ext, new object[] { newPath, (int)options, (int)version, null, err, warn });
                        ok = Convert.ToBoolean(res ?? false);
                    }
                }

                if (!ok)
                {
                    ErrorHandler.HandleError(proc, $"SaveAs failed for '{newPath}' (err={err}, warn={warn})");
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, $"Exception SaveAs: {newPath}", ex);
                return false;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        public static bool Close(ISldWorks app, IModelDoc2 doc)
        {
            return SolidWorksApiWrapper.CloseDocument(app, doc);
        }

        public static bool CloseByPath(ISldWorks app, string pathOrTitle)
        {
            const string proc = "SolidWorksFileOps.CloseByPath";
            ErrorHandler.PushCallStack(proc);
            try
            {
                var doc = TryGetOpenDoc(app, pathOrTitle);
                if (doc != null) return Close(app, doc);

                // Fallback: try closing by title if open but not resolved above
                if (app == null || string.IsNullOrWhiteSpace(pathOrTitle)) return false;
                var title = Path.GetFileName(pathOrTitle);
                if (!string.IsNullOrWhiteSpace(title))
                {
                    app.CloseDoc(title);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, $"Exception closing: {pathOrTitle}", ex);
                return false;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        public static bool Activate(ISldWorks app, string pathOrTitle)
        {
            const string proc = "SolidWorksFileOps.Activate";
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (app == null) return false;
                int errs = 0;
                var title = Path.GetFileName(pathOrTitle);
                if (string.IsNullOrWhiteSpace(title)) title = pathOrTitle;
                var resObj = app.ActivateDoc3(title, true, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref errs);
                int res = 0; try { res = Convert.ToInt32(resObj); } catch { }
                if (res != 1 || errs != 0)
                {
                    ErrorHandler.HandleError(proc, $"Activate failed for '{title}' (res={res}, err={errs})", null, "Warning");
                }
                return app.ActiveDoc != null;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, $"Exception activating: {pathOrTitle}", ex);
                return false;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        public static bool IsDirty(IModelDoc2 doc)
        {
            try { return doc != null && doc.GetSaveFlag(); }
            catch { return false; }
        }

        public static swDocumentTypes_e DetermineDocTypeByPath(string path)
        {
            var ext = (Path.GetExtension(path) ?? string.Empty).ToLowerInvariant();
            switch (ext)
            {
                case ".sldprt": return swDocumentTypes_e.swDocPART;
                case ".sldasm": return swDocumentTypes_e.swDocASSEMBLY;
                case ".slddrw": return swDocumentTypes_e.swDocDRAWING;
                default: return swDocumentTypes_e.swDocNONE;
            }
        }
    }
}
