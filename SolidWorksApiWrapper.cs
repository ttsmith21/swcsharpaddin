using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using NM.Core;
using System;

namespace NM.SwAddin
{
    #region Enums
    /// <summary>
    /// Document types (matches SolidWorks API swDocumentTypes_e)
    /// </summary>
    public enum SwDocumentTypes
    {
        SwDocNone = 0,
        SwDocPart = 1,
        SwDocAssembly = 2,
        SwDocDrawing = 3
    }

    /// <summary>
    /// Open doc options (matches swOpenDocOptions_e)
    /// </summary>
    public enum SwOpenDocOptions
    {
        Default = 0,
        Silent = 1,
        ReadOnly = 2,
        LoadLightweight = 4
    }

    /// <summary>
    /// Errors when opening a doc (matches swFileLoadError_e)
    /// </summary>
    public enum SwFileLoadError
    {
        SwGenericError = 1,
        SwFileNotFoundError = 2,
        SwFileRequiresRepairError = 3,
        SwFileLoadWarning = 5,
        SwFileLoadError_AlreadyOpen = 8
    }

    /// <summary>
    /// Errors when saving a doc (matches swFileSaveError_e)
    /// </summary>
    public enum SwFileSaveError
    {
        SwGenericSaveError = 1,
        SwFileSaveError_DiskFull = 2,
        SwFileSaveError_ReadOnly = 3,
        SwFileSaveError_AlreadyOpen = 4,
        SwFileSaveError_ReplaceFailed = 6,
        SwFileSaveError_ReferencedDocumentNotFound = 7
    }

    /// <summary>
    /// Save As options (matches swSaveAsOptions_e)
    /// </summary>
    public enum SwSaveAsOptions
    {
        SwSaveAsOptions_None = 0,
        SwSaveAsOptions_Silent = 1,
        SwSaveAsOptions_Copy = 2
    }

    /// <summary>
    /// Save As version options (matches swSaveAsVersion_e)
    /// </summary>
    public enum SwSaveAsVersion
    {
        SwSaveAsCurrentVersion = 0
    }

    /// <summary>
    /// Rebuild options (custom for ForceRebuildDoc)
    /// </summary>
    public enum SwRebuildOptions
    {
        SwRebuildActiveOnly = 0,
        SwForceRebuildAll = 1
    }

    /// <summary>
    /// Custom properties data types (matches swCustomInfoType_e)
    /// </summary>
    public enum SwCustomInfoType
    {
        SwCustomInfoText = 30,
        SwCustomInfoDate = 64,
        SwCustomInfoNumber = 3
    }

    /// <summary>
    /// Display modes (matches swViewDisplayMode_e)
    /// </summary>
    public enum SwViewDisplayMode
    {
        SwViewDisplayMode_Shaded = 1,
        SwViewDisplayMode_HiddenLinesRemoved = 2,
        SwViewDisplayMode_HiddenLinesGray = 3,
        SwViewDisplayMode_Wireframe = 4
    }
    #endregion

    /// <summary>
    /// SolidWorks API wrapper with error handling and validation.
    /// This class provides a "thin glue" layer for interacting with SolidWorks COM objects.
    /// </summary>
    public static class SolidWorksApiWrapper
    {
        #region Constants
        private const string ErrInvalidModel = "Invalid model reference";
        private const string ErrInvalidPart = "Invalid part document reference";
        private const string ErrInvalidName = "Invalid name parameter";
        private const string ErrInvalidApp = "Invalid SolidWorks application reference";
        private const string ErrInvalidValue = "Invalid parameter value";
        #endregion

        #region Validation Methods
        /// <summary>
        /// Validates that a model document is not null and is writable.
        /// </summary>
        /// <param name="swModel">The SolidWorks model to validate.</param>
        /// <param name="procedureName">The name of the calling procedure for error context.</param>
        /// <returns>True if the model is valid and writable, false otherwise.</returns>
        public static bool ValidateModel(IModelDoc2 swModel, string procedureName = "Unknown")
        {
            ErrorHandler.PushCallStack(procedureName);
            try
            {
                if (swModel == null)
                {
                    ErrorHandler.HandleError(procedureName, ErrInvalidModel);
                    return false;
                }

                if (swModel.IsOpenedReadOnly())
                {
                    ErrorHandler.HandleError(procedureName, "Document is read-only");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procedureName, "Model validation failed", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Validates that a string parameter is not null, empty, or whitespace.
        /// </summary>
        /// <param name="value">The string to validate.</param>
        /// <param name="procedureName">The name of the calling procedure for error context.</param>
        /// <param name="paramName">The name of the parameter for the error message.</param>
        /// <returns>True if the string is valid, false otherwise.</returns>
        public static bool ValidateString(string value, string procedureName, string paramName = "parameter")
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                ErrorHandler.HandleError(procedureName, $"Invalid {paramName}");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Validates that a double value falls within a specified range.
        /// </summary>
        /// <param name="value">The value to validate.</param>
        /// <param name="min">The minimum allowed value (inclusive).</param>
        /// <param name="max">The maximum allowed value (inclusive).</param>
        /// <param name="procedureName">The name of the calling procedure for error context.</param>
        /// <param name="paramName">The name of the parameter for the error message.</param>
        /// <returns>True if the value is in range, false otherwise.</returns>
        public static bool ValidateDoubleRange(double value, double min, double max, string procedureName, string paramName = "value")
        {
            if (value < min || value > max)
            {
                ErrorHandler.HandleError(procedureName, $"{paramName} out of range [{min}..{max}]");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Validates that a configuration exists in the specified model.
        /// </summary>
        /// <param name="swModel">The model to check.</param>
        /// <param name="configName">The name of the configuration to validate.</param>
        /// <returns>True if the configuration exists, false if not found.</returns>
        public static bool ValidateConfiguration(IModelDoc2 swModel, string configName)
        {
            const string procName = "ValidateConfiguration";
            if (!ValidateModel(swModel, procName)) return false;

            try
            {
                var swConfigMgr = swModel.ConfigurationManager;
                // Try to show the configuration - if it works, it exists
                string currentConfig = swConfigMgr.ActiveConfiguration.Name;
                bool exists = swModel.ShowConfiguration2(configName);
                if (exists && currentConfig != configName)
                {
                    // Restore original configuration
                    swModel.ShowConfiguration2(currentConfig);
                }
                return exists;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region Core Application Methods
        /// <summary>
        /// Gets the currently active SolidWorks document.
        /// </summary>
        /// <param name="swApp">The SolidWorks application object.</param>
        /// <returns>The active ModelDoc2, or null if none is active or an error occurs.</returns>
        public static IModelDoc2 GetActiveDoc(ISldWorks swApp)
        {
            const string procName = "GetActiveDoc";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (swApp == null)
                {
                    ErrorHandler.HandleError(procName, ErrInvalidApp);
                    return null;
                }
                return swApp.ActiveDoc as IModelDoc2;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Failed to get active document.", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Controls the visibility of the SolidWorks application window.
        /// </summary>
        /// <param name="swApp">The SolidWorks application object.</param>
        /// <param name="visible">True to show the application, false to hide it.</param>
        public static void SetSWVisibility(ISldWorks swApp, bool visible)
        {
            const string procName = "SetSWVisibility";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (swApp == null)
                {
                    ErrorHandler.HandleError(procName, ErrInvalidApp);
                    return;
                }
                swApp.Visible = visible;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Unable to set SolidWorks visibility.", ex);
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }
        #endregion

        #region Geometry Methods
        /// <summary>
        /// Gets the first solid body from a part document.
        /// </summary>
        /// <param name="swModel">The SolidWorks model to analyze.</param>
        /// <returns>The first solid body, or null if not found or an error occurs.</returns>
        public static IBody2 GetMainBody(IModelDoc2 swModel)
        {
            const string procName = "GetMainBody";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return null;

                var swPart = swModel as IPartDoc;
                if (swPart == null)
                {
                    ErrorHandler.HandleError(procName, "Model is not a part document.");
                    return null;
                }

                var vBodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
                if (vBodies == null || vBodies.Length == 0)
                {
                    ErrorHandler.HandleError(procName, "Part contains no solid bodies.");
                    return null;
                }

                return vBodies[0] as IBody2;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Error getting main body.", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Gets the largest planar face from a given body.
        /// </summary>
        /// <param name="swBody">The SolidWorks body to analyze.</param>
        /// <returns>The largest planar face, or null if not found or an error occurs.</returns>
        public static IFace2 GetLargestPlanarFace(IBody2 swBody)
        {
            const string procName = "GetLargestPlanarFace";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (swBody == null)
                {
                    ErrorHandler.HandleError(procName, "Input body is null.");
                    return null;
                }

                var vFaces = swBody.GetFaces() as object[];
                if (vFaces == null || vFaces.Length == 0)
                {
                    ErrorHandler.HandleError(procName, "Body contains no faces.");
                    return null;
                }

                IFace2 largestFace = null;
                double maxArea = 0;

                foreach (object face in vFaces)
                {
                    IFace2 swFace = face as IFace2;
                    if (swFace == null) continue;

                    ISurface swSurf = swFace.GetSurface() as ISurface;
                    if (swSurf != null && swSurf.IsPlane())
                    {
                        double currentArea = swFace.GetArea();
                        if (currentArea > maxArea)
                        {
                            maxArea = currentArea;
                            largestFace = swFace;
                        }
                    }
                }

                if (largestFace == null)
                {
                    ErrorHandler.HandleError(procName, "No planar faces found in body.", null, ErrorHandler.LogLevel.Warning);
                }

                return largestFace;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Error getting largest planar face.", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Checks if the model contains a SheetMetal or FlatPattern feature.
        /// VBA approach: Check body.GetFeatures() first, then fall back to model traversal.
        /// </summary>
        /// <param name="swModel">The model to check.</param>
        /// <returns>True if a sheet metal feature is found, false otherwise.</returns>
        public static bool HasSheetMetalFeature(IModelDoc2 swModel)
        {
            const string procName = "HasSheetMetalFeature";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;

                // VBA approach: Check body features first (swBody.GetFeatures)
                // This is how VBA's SMInsertBends() checks: Features = swBody.GetFeatures
                var body = GetMainBody(swModel);
                if (body != null)
                {
                    var bodyFeatures = body.GetFeatures() as object[];
                    if (bodyFeatures != null)
                    {
                        foreach (var featObj in bodyFeatures)
                        {
                            var feat = featObj as IFeature;
                            if (feat == null) continue;
                            string featType = feat.GetTypeName2();
                            // VBA checks: If Feature = "SheetMetal" Then (exact match)
                            if (string.Equals(featType, "SheetMetal", StringComparison.OrdinalIgnoreCase))
                            {
                                return true;
                            }
                        }
                    }
                }

                // Fallback: also check model feature tree for FlatPattern (may not be on body)
                var swFeat = swModel.FirstFeature() as IFeature;
                while (swFeat != null)
                {
                    string featType = swFeat.GetTypeName2();
                    if (featType.IndexOf("SheetMetal", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        featType.IndexOf("FlatPattern", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }
                    swFeat = swFeat.GetNextFeature() as IFeature;
                }
                return false;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Error checking for sheet metal features.", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Gets the longest linear edge from a given body.
        /// </summary>
        /// <param name="swBody">The SolidWorks body to analyze.</param>
        /// <returns>The longest linear edge, or null if not found or an error occurs.</returns>
        public static IEdge GetLongestLinearEdge(IBody2 swBody)
        {
            const string procName = "GetLongestLinearEdge";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (swBody == null)
                {
                    ErrorHandler.HandleError(procName, "Input body is null.");
                    return null;
                }

                var vEdges = swBody.GetEdges() as object[];
                if (vEdges == null || vEdges.Length == 0)
                {
                    ErrorHandler.HandleError(procName, "Body contains no edges.");
                    return null;
                }

                IEdge longestEdge = null;
                double maxLength = 0;

                foreach (object edge in vEdges)
                {
                    IEdge swEdge = edge as IEdge;
                    if (swEdge == null) continue;

                    ICurve swCurve = swEdge.GetCurve() as ICurve;
                    if (swCurve != null && swCurve.IsLine())
                    {
                        var curveParams = swEdge.GetCurveParams2() as double[];
                        // In C#, array indices for GetCurveParams2 are 0-based.
                        // The start and end parameters are at indices 0 and 1 for a line.
                        // The VBA code was incorrect using 6 and 7.
                        double length = swCurve.GetLength2(curveParams[0], curveParams[1]);
                        if (length > maxLength)
                        {
                            maxLength = length;
                            longestEdge = swEdge;
                        }
                    }
                }

                if (longestEdge == null)
                {
                    ErrorHandler.HandleError(procName, "No linear edges found in body.", null, ErrorHandler.LogLevel.Warning);
                }

                return longestEdge;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Error getting longest linear edge.", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }
        #endregion

        #region User Preferences Methods
        /// <summary>
        /// Retrieves a boolean preference from SolidWorks.
        /// </summary>
        /// <param name="swApp">The SolidWorks application object.</param>
        /// <param name="prefType">The preference type constant.</param>
        /// <returns>The preference value, or false if an error occurs.</returns>
        public static bool GetSWUserPreferenceToggle(ISldWorks swApp, int prefType)
        {
            const string procName = "GetSWUserPreferenceToggle";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (swApp == null)
                {
                    ErrorHandler.HandleError(procName, ErrInvalidApp);
                    return false;
                }
                return swApp.GetUserPreferenceToggle(prefType);
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Failed to get preference toggle. Type={prefType}", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Sets a boolean preference in the SolidWorks application.
        /// </summary>
        /// <param name="swApp">The SolidWorks application object.</param>
        /// <param name="prefType">The preference type constant.</param>
        /// <param name="state">True to enable, false to disable the preference.</param>
        public static void SetSWUserPreferenceToggle(ISldWorks swApp, int prefType, bool state)
        {
            const string procName = "SetSWUserPreferenceToggle";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (swApp == null)
                {
                    ErrorHandler.HandleError(procName, ErrInvalidApp);
                    return;
                }
                swApp.SetUserPreferenceToggle(prefType, state);
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Failed to set preference toggle. Type={prefType}", ex);
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }
        #endregion

        #region Document Management Methods
        /// <summary>
        /// Closes all open documents in the current SolidWorks session.
        /// </summary>
        /// <param name="swApp">The SolidWorks application object.</param>
        /// <returns>True if all documents closed successfully, false otherwise.</returns>
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

                // Close documents in reverse order
                for (int i = docCount - 1; i >= 0; i--)
                {
                    if (swDocs[i] is IModelDoc2 swModel)
                    {
                        string docTitle = swModel.GetTitle();
                        swApp.CloseDoc(docTitle);
                    }
                }

                // Verify all closed
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
        /// <param name="swModel">The document to rebuild.</param>
        /// <param name="rebuildOption">The rebuild option (active or all configs).</param>
        /// <returns>True if rebuild successful or not needed, false if error occurs.</returns>
        public static bool ForceRebuildDoc(IModelDoc2 swModel, SwRebuildOptions rebuildOption = SwRebuildOptions.SwRebuildActiveOnly)
        {
            const string procName = "ForceRebuildDoc";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;

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
        #endregion

        #region Selection Methods
        /// <summary>
        /// Clears all selected entities in the specified model.
        /// </summary>
        /// <param name="swModel">The model to clear selection in.</param>
        /// <returns>True if selection cleared successfully, false otherwise.</returns>
        public static bool ClearSelection(IModelDoc2 swModel)
        {
            const string procName = "ClearSelection";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;

                swModel.ClearSelection2(true);
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Failed to clear selection in: {swModel.GetTitle()}", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Selects a SolidWorks entity by name and type at specified coordinates.
        /// </summary>
        /// <param name="swModel">The model to select in.</param>
        /// <param name="objName">The name of object to select.</param>
        /// <param name="objType">The type of object (e.g. "PLANE", "EDGE", etc).</param>
        /// <param name="x">X coordinate for selection point.</param>
        /// <param name="y">Y coordinate for selection point.</param>
        /// <param name="z">Z coordinate for selection point.</param>
        /// <param name="appendSelection">True to append to current selection.</param>
        /// <param name="mark">Selection mark.</param>
        /// <returns>True if selection successful, false if fails.</returns>
        public static bool SelectByName(IModelDoc2 swModel, string objName, string objType, 
            double x, double y, double z, bool appendSelection = false, int mark = 0)
        {
            const string procName = "SelectByName";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;
                if (!ValidateString(objName, procName, "object name")) return false;
                if (!ValidateString(objType, procName, "object type")) return false;

                var swExt = swModel.Extension;
                bool result = swExt.SelectByID2(objName, objType, x, y, z, appendSelection, mark, null, 0);

                if (!result)
                {
                    ErrorHandler.HandleError(procName, $"Failed to select: [{objName}] Type: {objType}", null, ErrorHandler.LogLevel.Warning);
                }

                return result;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception selecting: {objName}", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }
        #endregion

        #region Feature Methods
        /// <summary>
        /// Retrieves a SolidWorks feature by its name.
        /// </summary>
        /// <param name="swModel">The model containing the feature.</param>
        /// <param name="featName">The name of feature to retrieve.</param>
        /// <returns>The feature object if found, null if not found or error occurs.</returns>
        public static IFeature GetFeatureByName(IModelDoc2 swModel, string featName)
        {
            const string procName = "GetFeatureByName";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return null;
                if (!ValidateString(featName, procName, "feature name")) return null;

                var swFeat = FindFeatureByName(swModel, featName);
                if (swFeat == null)
                {
                    ErrorHandler.HandleError(procName, $"Feature not found: {featName}", null, ErrorHandler.LogLevel.Warning);
                }

                return swFeat;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Failed to get feature: {featName}", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Suppresses a SolidWorks feature by name.
        /// </summary>
        /// <param name="swModel">The model containing the feature.</param>
        /// <param name="featName">The name of feature to suppress.</param>
        /// <returns>True if feature successfully suppressed, false if error occurs.</returns>
        public static bool SuppressFeature(IModelDoc2 swModel, string featName)
        {
            const string procName = "SuppressFeature";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;
                if (!ValidateString(featName, procName, "feature name")) return false;

                var swFeat = FindFeatureByName(swModel, featName);
                if (swFeat == null)
                {
                    ErrorHandler.HandleError(procName, $"Feature not found: {featName}", null, ErrorHandler.LogLevel.Warning);
                    return false;
                }

                // 0 = suppressed, 2 = standard option
                object ret = swFeat.SetSuppression2(0, 2, null);
                bool success = Convert.ToBoolean(ret);

                if (!success)
                {
                    ErrorHandler.HandleError(procName, $"Failed to suppress: {featName}", null, ErrorHandler.LogLevel.Warning);
                }

                return success;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception suppressing: {featName}", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Unsuppresses a SolidWorks feature by name.
        /// </summary>
        /// <param name="swModel">The model containing the feature.</param>
        /// <param name="featName">The name of feature to unsuppress.</param>
        /// <returns>True if feature successfully unsuppressed, false if error occurs.</returns>
        public static bool UnsuppressFeature(IModelDoc2 swModel, string featName)
        {
            const string procName = "UnsuppressFeature";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;
                if (!ValidateString(featName, procName, "feature name")) return false;

                var swFeat = FindFeatureByName(swModel, featName);
                if (swFeat == null)
                {
                    ErrorHandler.HandleError(procName, $"Feature not found: {featName}", null, ErrorHandler.LogLevel.Warning);
                    return false;
                }

                // 1 = unsuppressed, 2 = standard option
                object ret = swFeat.SetSuppression2(1, 2, null);
                bool success = Convert.ToBoolean(ret);

                if (!success)
                {
                    ErrorHandler.HandleError(procName, $"Failed to unsuppress: {featName}", null, ErrorHandler.LogLevel.Warning);
                }

                return success;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception unsuppressing: {featName}", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Renames a SolidWorks feature from old name to new name.
        /// </summary>
        /// <param name="swModel">The model containing the feature.</param>
        /// <param name="oldName">The current name of feature.</param>
        /// <param name="newName">The new name for feature.</param>
        /// <returns>True if feature successfully renamed, false if error occurs.</returns>
        public static bool RenameFeature(IModelDoc2 swModel, string oldName, string newName)
        {
            const string procName = "RenameFeature";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;
                if (!ValidateString(oldName, procName, "old name")) return false;
                if (!ValidateString(newName, procName, "new name")) return false;

                var swFeat = FindFeatureByName(swModel, oldName);
                if (swFeat == null)
                {
                    ErrorHandler.HandleError(procName, $"Feature not found: {oldName}", null, ErrorHandler.LogLevel.Warning);
                    return false;
                }

                swFeat.Name = newName;
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Failed to rename: {oldName} to {newName}", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }
        #endregion

        #region Configuration Methods
        /// <summary>
        /// Retrieves the ConfigurationManager object from a SolidWorks model.
        /// </summary>
        /// <param name="swModel">The model to get configuration manager from.</param>
        /// <returns>The configuration manager object or null if error.</returns>
        public static IConfigurationManager GetConfigurationManager(IModelDoc2 swModel)
        {
            const string procName = "GetConfigurationManager";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return null;
                return swModel.ConfigurationManager;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Failed to get configuration manager.", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Retrieves the active configuration from a SolidWorks model.
        /// </summary>
        /// <param name="swModel">The model to get configuration from.</param>
        /// <returns>The active configuration object or null if error occurs.</returns>
        public static IConfiguration GetActiveConfiguration(IModelDoc2 swModel)
        {
            const string procName = "GetActiveConfiguration";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return null;

                var swCfgMgr = swModel.ConfigurationManager;
                return swCfgMgr.ActiveConfiguration as IConfiguration;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Failed to get active configuration.", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Activates a specified configuration in a SolidWorks model.
        /// </summary>
        /// <param name="swModel">The model containing the configuration.</param>
        /// <param name="configName">The name of the configuration to activate.</param>
        /// <returns>True if configuration is active (set or already active), false if error.</returns>
        public static bool SetActiveConfiguration(IModelDoc2 swModel, string configName)
        {
            const string procName = "SetActiveConfiguration";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;
                if (!ValidateString(configName, procName, "configuration name")) return false;

                var swCfgMgr = swModel.ConfigurationManager;
                if (swCfgMgr == null)
                {
                    ErrorHandler.HandleError(procName, "Failed to get ConfigurationManager");
                    return false;
                }

                ErrorHandler.DebugLog($"Attempting to activate configuration: {configName}");

                // Check if the desired configuration is already active
                if (swCfgMgr.ActiveConfiguration.Name != configName)
                {
                    // Only attempt to switch if it's not already active
                    if (!swModel.ShowConfiguration2(configName))
                    {
                        ErrorHandler.HandleError(procName, $"Failed to set configuration: {configName}");
                        return false;
                    }
                }
                else
                {
                    ErrorHandler.DebugLog($"Configuration '{configName}' is already active in {swModel.GetPathName()}");
                }

                // Verify activation
                bool success = swCfgMgr.ActiveConfiguration.Name == configName;
                if (!success)
                {
                    ErrorHandler.HandleError(procName, 
                        $"Verification failed: Active configuration is '{swCfgMgr.ActiveConfiguration.Name}' not '{configName}'", 
                        null, ErrorHandler.LogLevel.Warning);
                }
                else
                {
                    ErrorHandler.DebugLog($"Configuration activated: {configName}");
                }

                return success;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Error activating configuration: {configName}", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Creates a new configuration in a SolidWorks model.
        /// </summary>
        /// <param name="swModel">The model to add configuration to.</param>
        /// <param name="newConfigName">The name for new configuration.</param>
        /// <param name="comment">The configuration comment.</param>
        /// <param name="altName">The alternate name.</param>
        /// <returns>The name of created configuration or empty if failed.</returns>
        public static string AddConfiguration(IModelDoc2 swModel, string newConfigName, string comment = "", string altName = "")
        {
            const string procName = "AddConfiguration";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return "";
                if (!ValidateString(newConfigName, procName, "configuration name")) return "";

                ErrorHandler.DebugLog("Attempting AddConfiguration3 on ModelDoc2...");

                // Call AddConfiguration3 directly
                swModel.AddConfiguration3(newConfigName, altName, comment, 0);

                // Verify configuration was created by getting active config name
                string result = swModel.ConfigurationManager.ActiveConfiguration.Name;
                ErrorHandler.DebugLog($"Created configuration: {result}");

                swModel.ForceRebuild3(true);
                return result;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception creating configuration: {newConfigName}", ex);
                return "";
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Validates configuration name contains only legal characters.
        /// </summary>
        /// <param name="configName">The name to validate.</param>
        /// <returns>True if name is valid, false if contains illegal chars.</returns>
        public static bool IsValidConfigName(string configName)
        {
            const string illegalChars = "/\\:*?\"<>|";
            return !string.IsNullOrEmpty(configName) && 
                   configName.IndexOfAny(illegalChars.ToCharArray()) == -1;
        }
        #endregion

        #region Material Methods
        /// <summary>
        /// Sets the material for the active configuration of a SolidWorks part.
        /// </summary>
        /// <param name="swModel">Model to modify.</param>
        /// <param name="materialName">Name of material to apply.</param>
        /// <returns>True if material set successfully, false otherwise.</returns>
        public static bool SetMaterialName(IModelDoc2 swModel, string materialName)
        {
            const string procName = "SetMaterialName";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;
                if (!ValidateString(materialName, procName, "material name")) return false;

                var swPart = swModel as IPartDoc;
                if (swPart == null)
                {
                    ErrorHandler.HandleError(procName, ErrInvalidPart);
                    return false;
                }

                string configName = swModel.ConfigurationManager.ActiveConfiguration.Name;
                // Use configured materials database path if available
                string dbPath = NM.Core.Configuration.FilePaths.MaterialPropertyFilePath;
                if (string.IsNullOrWhiteSpace(dbPath)) dbPath = string.Empty;

                // Apply material to the active configuration
                swPart.SetMaterialPropertyName2(configName, dbPath, materialName);

                // Verify
                string verifyDb;
                string applied = swPart.GetMaterialPropertyName2(configName, out verifyDb);
                return string.Equals(applied, materialName, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Failed to set material: {materialName}", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Gets the material name from the active configuration of a SolidWorks part.
        /// </summary>
        /// <param name="swModel">Model to query.</param>
        /// <returns>Material name or empty string if not set or error occurs.</returns>
        public static string GetMaterialName(IModelDoc2 swModel)
        {
            const string procName = "GetMaterialName";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return string.Empty;

                var swPart = swModel as IPartDoc;
                if (swPart == null)
                {
                    ErrorHandler.HandleError(procName, ErrInvalidPart);
                    return string.Empty;
                }

                string configName = swModel.ConfigurationManager.ActiveConfiguration.Name;
                string dbName;
                string material = swPart.GetMaterialPropertyName2(configName, out dbName);

                if (NM.Core.Configuration.Logging.EnableDebugMode && !string.IsNullOrEmpty(dbName))
                {
                    ErrorHandler.HandleError(procName, $"Material retrieved from database: {dbName}", null, ErrorHandler.LogLevel.Info);
                }

                return material ?? string.Empty;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Failed to get material name", ex);
                return string.Empty;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }
        #endregion

        #region View / Display Methods
        /// <summary>
        /// Performs a Zoom to Fit on the active view of the given model.
        /// /// </summary>
        public static void ZoomToFit(IModelDoc2 swModel)
        {
            const string procName = "ZoomToFit";
            if (!ValidateModel(swModel, procName)) return;

            try
            {
                var swView = swModel.ActiveView as IModelView;
                // Use ModelDoc2 zoom-to-fit since IModelView.ZoomToFit is not available in this interop
                swModel.ViewZoomtofit2();
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Failed to ZoomToFit on model: {swModel.GetTitle()}", ex);
            }
        }

        /// <summary>
        /// Changes the display mode (wireframe, shaded, etc.) in the active view.
        /// </summary>
        public static bool SetViewDisplayMode(IModelDoc2 swModel, SwViewDisplayMode displayMode)
        {
            const string procName = "SetViewDisplayMode";
            if (!ValidateModel(swModel, procName)) return false;

            // Validate display mode enum value range
            if (displayMode < SwViewDisplayMode.SwViewDisplayMode_Shaded ||
                displayMode > SwViewDisplayMode.SwViewDisplayMode_Wireframe)
            {
                ErrorHandler.HandleError(procName, "Invalid display mode value");
                return false;
            }

            try
            {
                var swView = swModel.ActiveView as IModelView;
                if (swView == null)
                {
                    ErrorHandler.HandleError(procName, "No active view to set display mode.", null, ErrorHandler.LogLevel.Warning);
                    return false;
                }

                // Some interop versions use ModelDoc2/View display setting via SetDisplayMode
                swView = swModel.ActiveView as IModelView;
                swView.DisplayMode = (int)displayMode;
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Failed to set view display mode.", ex);
                return false;
            }
        }
        #endregion

        #region Mass Properties Methods
        /// <summary>
        /// Retrieves mass, volume, and center of mass from a model.
        /// </summary>
        public static bool GetAllMassProperties(IModelDoc2 swModel, out double mass, out double volume, out double comX, out double comY, out double comZ)
        {
            const string procName = "GetAllMassProperties";
            mass = 0; volume = 0; comX = 0; comY = 0; comZ = 0;

            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;

                var swExt = swModel.Extension;
                var swMassProp = swExt.CreateMassProperty() as IMassProperty;
                if (swMassProp == null)
                {
                    ErrorHandler.HandleError(procName, "Failed to create mass property object");
                    return false;
                }

                mass = swMassProp.Mass;
                volume = swMassProp.Volume;
                var arrCOM = swMassProp.CenterOfMass as object[];
                if (arrCOM == null || arrCOM.Length != 3)
                {
                    ErrorHandler.HandleError(procName, "Invalid center of mass data", null, ErrorHandler.LogLevel.Warning);
                    return false;
                }

                comX = Convert.ToDouble(arrCOM[0]);
                comY = Convert.ToDouble(arrCOM[1]);
                comZ = Convert.ToDouble(arrCOM[2]);
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Error getting mass properties", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Gets current model volume using mass properties.
        /// </summary>
        public static double GetModelVolume(IModelDoc2 swModel)
        {
            const string procName = "GetModelVolume";
            try
            {
                if (!ValidateModel(swModel, procName)) return -1;

                var swMassProp = swModel.Extension.CreateMassProperty() as IMassProperty;
                if (swMassProp == null)
                {
                    ErrorHandler.HandleError(procName, "Failed to create mass property", null, ErrorHandler.LogLevel.Warning);
                    return -1;
                }

                return Math.Round(swMassProp.Volume, 6);
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Exception getting model volume", ex);
                return -1;
            }
        }

        /// <summary>
        /// Gets current model mass using mass properties.
        /// </summary>
        public static double GetModelMass(IModelDoc2 swModel)
        {
            const string procName = "GetModelMass";
            try
            {
                if (!ValidateModel(swModel, procName)) return -1;

                var swMassProp = swModel.Extension.CreateMassProperty() as IMassProperty;
                if (swMassProp == null)
                {
                    ErrorHandler.HandleError(procName, "Failed to create mass property", null, ErrorHandler.LogLevel.Warning);
                    return -1;
                }

                return Math.Round(swMassProp.Mass, 6);
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Exception getting model mass", ex);
                return -1;
            }
        }
        #endregion

        #region Sketch Methods
        /// <summary>
        /// Validates that a SketchManager reference is not null.
        /// </summary>
        public static bool ValidateSketchManager(ISketchManager swSkMgr, string procedureName)
        {
            if (swSkMgr == null)
            {
                ErrorHandler.HandleError(procedureName, "Invalid SketchManager reference");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Retrieves the SketchManager from a SolidWorks model.
        /// </summary>
        public static ISketchManager GetSketchManager(IModelDoc2 swModel)
        {
            const string procName = "GetSketchManager";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return null;
                return swModel.SketchManager;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Failed to get sketch manager", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Creates a line in the active sketch between two points.
        /// </summary>
        public static ISketchSegment CreateSketchLine(IModelDoc2 swModel, double x1, double y1, double x2, double y2)
        {
            const string procName = "CreateSketchLine";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return null;
                if (x1 == x2 && y1 == y2)
                {
                    ErrorHandler.HandleError(procName, "Zero-length line not allowed", null, ErrorHandler.LogLevel.Warning);
                    return null;
                }

                var swSkMgr = swModel.SketchManager;
                if (!ValidateSketchManager(swSkMgr, procName)) return null;

                var seg = swSkMgr.CreateLine(x1, y1, 0.0, x2, y2, 0.0) as ISketchSegment;
                if (seg == null)
                {
                    ErrorHandler.HandleError(procName, "Failed to create line", null, ErrorHandler.LogLevel.Warning);
                }
                return seg;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Exception creating line", ex);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Creates a new sketch on the specified plane in a SolidWorks model.
        /// </summary>
        public static bool StartSketchOnPlane(IModelDoc2 swModel, string planeName)
        {
            const string procName = "StartSketchOnPlane";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;
                if (!ValidateString(planeName, procName, "plane name")) return false;

                if (!ClearSelection(swModel)) return false;

                if (!SelectByName(swModel, planeName, "PLANE", 0, 0, 0))
                {
                    ErrorHandler.HandleError(procName, $"Failed to select plane: {planeName}", null, ErrorHandler.LogLevel.Warning);
                    return false;
                }

                var swSkMgr = swModel.SketchManager;
                swSkMgr.InsertSketch(true);
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception creating sketch on: {planeName}", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Exits sketch editing mode in a SolidWorks model.
        /// </summary>
        public static bool EndSketch(IModelDoc2 swModel)
        {
            const string procName = "EndSketch";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;

                var swSkMgr = swModel.SketchManager;
                if (!ValidateSketchManager(swSkMgr, procName)) return false;

                swSkMgr.InsertSketch(true);
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, "Failed to exit sketch mode", ex);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }
        #endregion

        #region Custom Properties Methods
        /// <summary>
        /// Adds a custom property to a SolidWorks model.
        /// </summary>
        public static bool AddCustomProperty(IModelDoc2 swModel, string propName, swCustomInfoType_e propType, string propValue, string configName)
        {
            const string procName = "AddCustomProperty";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;
                if (!ValidateString(propName, procName, "property name")) return false;

                var mgr = swModel.Extension.get_CustomPropertyManager(configName);
                int addResult = mgr.Add3(propName, (int)propType, propValue, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);

                // Verify by querying back since return codes vary by version/options
                bool exists = TryGetProperty(mgr, propName, out var currentValue);
                bool ok = exists; // we only require it to exist; value may be evaluated text

                if (!ok)
                {
                    // Fallback attempts
                    int setResult = mgr.Set2(propName, propValue);
                    exists = TryGetProperty(mgr, propName, out currentValue);
                    if (!exists)
                    {
                        try
                        {
                            int add2Result = mgr.Add2(propName, (int)propType, propValue);
                        }
                        catch { }
                        exists = TryGetProperty(mgr, propName, out currentValue);
                    }

                    ErrorHandler.DebugLog($"AddCustomProperty failed (code={addResult}) for '{propName}'.");
                }
                return ok || exists;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception: {ex.Message}", ex, ErrorHandler.LogLevel.Error, $"Property: {propName}");
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Sets the value of an existing custom property.
        /// </summary>
        public static bool SetCustomProperty(IModelDoc2 swModel, string propName, string propValue, string configName)
        {
            const string procName = "SetCustomProperty";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;
                if (!ValidateString(propName, procName, "property name")) return false;

                var mgr = swModel.Extension.get_CustomPropertyManager(configName);
                int setResult = mgr.Set2(propName, propValue);

                // Verify
                bool exists = TryGetProperty(mgr, propName, out var currentValue);
                if (!exists)
                {
                    // If Set2 failed due to not existing, try Add2
                    try { mgr.Add2(propName, (int)swCustomInfoType_e.swCustomInfoText, propValue); } catch { }
                    exists = TryGetProperty(mgr, propName, out currentValue);
                }
                return exists;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception: {ex.Message}", ex, ErrorHandler.LogLevel.Error, $"Property: {propName}");
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Retrieves all custom properties from a SolidWorks model.
        /// </summary>
        public static bool GetCustomProperties(IModelDoc2 swModel, string configName, out string[] propNames, out int[] propTypes, out string[] propValues)
        {
            const string procName = "GetCustomProperties";
            propNames = Array.Empty<string>();
            propTypes = Array.Empty<int>();
            propValues = Array.Empty<string>();

            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;

                var mgr = swModel.Extension.get_CustomPropertyManager(configName);
                object namesObj = null, typesObj = null, valuesObj = null, resolved = null, links = null;
                mgr.GetAll3(ref namesObj, ref typesObj, ref valuesObj, ref resolved, ref links);

                var namesArr = namesObj as object[] ?? namesObj as string[];
                var typesArr = typesObj as object[];
                var valuesArr = valuesObj as object[] ?? valuesObj as string[];

                if (namesArr == null || typesArr == null || valuesArr == null)
                {
                    return true; // no properties
                }

                int len = Math.Min(namesArr.Length, Math.Min(typesArr.Length, valuesArr.Length));
                propNames = new string[len];
                propTypes = new int[len];
                propValues = new string[len];
                for (int i = 0; i < len; i++)
                {
                    object n = (namesArr is string[]) ? ((string[])namesArr)[i] : namesArr[i];
                    object v = (valuesArr is string[]) ? ((string[])valuesArr)[i] : valuesArr[i];
                    propNames[i] = n?.ToString() ?? string.Empty;
                    propTypes[i] = Convert.ToInt32(typesArr[i] ?? 0);
                    propValues[i] = v?.ToString() ?? string.Empty;
                }
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception: {ex.Message}", ex, ErrorHandler.LogLevel.Error);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Deletes a custom property from a SolidWorks model.
        /// </summary>
        public static bool DeleteCustomProperty(IModelDoc2 swModel, string propName, string configName)
        {
            const string procName = "DeleteCustomProperty";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!ValidateModel(swModel, procName)) return false;
                if (!ValidateString(propName, procName, "property name")) return false;

                var mgr = swModel.Extension.get_CustomPropertyManager(configName);
                int delResult = mgr.Delete2(propName);

                bool exists = TryGetProperty(mgr, propName, out var _);
                bool ok = !exists;
                if (!ok)
                {
                    ErrorHandler.DebugLog($"DeleteCustomProperty failed (code={delResult}) for '{propName}'.");
                }
                return ok;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception: {ex.Message}", ex, ErrorHandler.LogLevel.Error, $"Property: {propName}");
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }
        #endregion

        #region Document IO Methods
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
                if (!ValidateModel(swModel, procName)) return false;
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
        #endregion

        #region Helper Methods
        private static IFeature FindFeatureByName(IModelDoc2 swModel, string featName)
        {
            if (swModel == null || string.IsNullOrWhiteSpace(featName)) return null;

            IFeature feat = swModel.FirstFeature() as IFeature;
            while (feat != null)
            {
                try
                {
                    if (string.Equals(feat.Name, featName, StringComparison.OrdinalIgnoreCase))
                    {
                        return feat;
                    }
                }
                catch
                {
                    // ignore and continue
                }
                feat = feat.GetNextFeature() as IFeature;
            }
            return null;
        }

        private static bool TryGetProperty(ICustomPropertyManager mgr, string propName, out string val)
        {
            val = string.Empty;
            try
            {
                // Use Get5 if available; fallback to Get2 signature
                // Here we rely on Get5 via dynamic cast; else use Get2
                // But to remain strongly typed across interop versions, use GetAll3
                object nObj = null, tObj = null, vObj = null, rObj = null, lObj = null;
                mgr.GetAll3(ref nObj, ref tObj, ref vObj, ref rObj, ref lObj);
                var names = nObj as object[] ?? nObj as string[];
                var values = vObj as object[] ?? vObj as string[];
                if (names == null || values == null) return false;
                for (int i = 0; i < Math.Min(names.Length, values.Length); i++)
                {
                    var n = (names is string[]) ? ((string[])names)[i] : names[i]?.ToString();
                    if (string.Equals(n, propName, StringComparison.OrdinalIgnoreCase))
                    {
                        val = (values is string[]) ? ((string[])values)[i] : values[i]?.ToString();
                        return true;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region Custom Property Helpers
        /// <summary>
        /// Gets a custom property value from a model.
        /// </summary>
        public static string GetCustomPropertyValue(IModelDoc2 model, string propName, string configName = "")
        {
            if (model == null || string.IsNullOrEmpty(propName)) return string.Empty;
            try
            {
                var ext = model.Extension;
                if (ext == null) return string.Empty;

                var mgr = string.IsNullOrEmpty(configName)
                    ? ext.get_CustomPropertyManager(string.Empty)
                    : ext.get_CustomPropertyManager(configName);

                if (mgr == null) return string.Empty;

                string val = string.Empty;
                string resolved = string.Empty;
                bool wasResolved = false;
                mgr.Get5(propName, false, out val, out resolved, out wasResolved);
                return resolved ?? val ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Gets the fixed face from a model (for sheet metal processing).
        /// Returns null if not found.
        /// </summary>
        public static IFace2 GetFixedFace(IModelDoc2 model)
        {
            if (model == null) return null;
            try
            {
                var part = model as IPartDoc;
                if (part == null) return null;

                var bodies = part.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
                if (bodies == null || bodies.Length == 0) return null;

                foreach (var b in bodies)
                {
                    var body = b as IBody2;
                    if (body == null) continue;
                    var faces = body.GetFaces() as object[];
                    if (faces == null) continue;
                    foreach (var f in faces)
                    {
                        var face = f as IFace2;
                        if (face == null) continue;
                        var surf = face.IGetSurface();
                        if (surf != null && surf.IsPlane())
                        {
                            // Return first planar face as fallback
                            return face;
                        }
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
        #endregion
    }
}
