using System;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin
{
    /// <summary>
    /// Document lifecycle operations (open, save, close, rebuild) for SolidWorks.
    /// Extracted from SolidWorksApiWrapper for single-responsibility.
    /// </summary>
    public static class SwDocumentHelper
    {
        private const string ErrInvalidApp = "Invalid SolidWorks application reference";

        /// <summary>
        /// Opens a SolidWorks document by path using OpenDoc6.
        /// </summary>
        public static IModelDoc2 OpenDocument(ISldWorks swApp, string filePath, swDocumentTypes_e docType, swOpenDocOptions_e options, out int err, out int warn)
        {
            const string procName = "OpenDocument";
            ErrorHandler.PushCallStack(procName);
            err = 0; warn = 0;
            try
            {
                if (swApp == null)
                {
                    ErrorHandler.HandleError(procName, ErrInvalidApp);
                    return null;
                }
                if (string.IsNullOrWhiteSpace(filePath))
                {
                    ErrorHandler.HandleError(procName, "File path is empty");
                    return null;
                }
                var model = swApp.OpenDoc6(filePath, (int)docType, (int)options, "", ref err, ref warn) as IModelDoc2;
                if (model == null || err != 0)
                {
                    ErrorHandler.HandleError(procName, $"Failed to open: {filePath} (err={err}, warn={warn})");
                }
                return model;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception opening: {filePath}", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Saves the specified SolidWorks document using Save3.
        /// </summary>
        public static bool SaveDocument(IModelDoc2 swModel, swSaveAsOptions_e options = swSaveAsOptions_e.swSaveAsOptions_Silent)
        {
            const string procName = "SaveDocument";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return false;
                int err = 0, warn = 0;
                bool ok = swModel.Save3((int)options, ref err, ref warn);
                if (!ok || err != 0)
                {
                    ErrorHandler.HandleError(procName, $"Save failed for '{swModel.GetTitle()}' (err={err}, warn={warn})");
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Exception saving document", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Closes the specified document via the application.
        /// </summary>
        public static bool CloseDocument(ISldWorks swApp, IModelDoc2 swModel)
        {
            const string procName = "CloseDocument";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (swApp == null)
                {
                    ErrorHandler.HandleError(procName, ErrInvalidApp);
                    return false;
                }
                if (swModel == null) return true;
                string title = swModel.GetTitle();
                swApp.CloseDoc(title);
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Exception closing document", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Closes all open documents in the current SolidWorks session.
        /// </summary>
        public static bool CloseAllDocuments(ISldWorks swApp)
        {
            const string procName = "CloseAllDocuments";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (swApp == null)
                {
                    ErrorHandler.HandleError(procName, ErrInvalidApp);
                    return false;
                }

                int docCount = swApp.GetDocumentCount();
                if (docCount == 0) return true;

                var swDocs = swApp.GetDocuments() as object[];
                if (swDocs == null) return true;

                for (int i = docCount - 1; i >= 0; i--)
                {
                    if (swDocs[i] is IModelDoc2 swModel)
                    {
                        string docTitle = swModel.GetTitle();
                        swApp.CloseDoc(docTitle);
                    }
                }

                return swApp.GetDocumentCount() == 0;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Failed to close all documents.", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Forces a rebuild of the specified SolidWorks document if needed.
        /// </summary>
        public static bool ForceRebuildDoc(IModelDoc2 swModel, SwRebuildOptions rebuildOption = SwRebuildOptions.SwRebuildActiveOnly)
        {
            const string procName = "ForceRebuildDoc";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return false;

                string docTitle = swModel.GetTitle();
                int rebuildStatus = swModel.Extension.NeedsRebuild2;

                ErrorHandler.DebugLog($"Model '{docTitle}' rebuild status: {rebuildStatus} (0=Fully Rebuilt, 1=Non-Frozen Needs Rebuild, 2=Frozen Needs Rebuild)");

                if (rebuildStatus != 0)
                {
                    bool rebuildAll = rebuildOption == SwRebuildOptions.SwForceRebuildAll;
                    if (!swModel.ForceRebuild3(rebuildAll))
                    {
                        ErrorHandler.HandleError(procName, $"Failed to rebuild {(rebuildAll ? "all" : "active")} in '{docTitle}'. Status: {rebuildStatus}", null, ErrorHandler.LogLevel.Warning);
                        return false;
                    }
                }
                else
                {
                    ErrorHandler.DebugLog($"Model '{docTitle}' is fully rebuilt; no rebuild needed");
                }

                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Failed to rebuild document.", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }
    }
}
