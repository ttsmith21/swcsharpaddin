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

        // Document Management Methods moved to SwDocumentHelper.cs

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

        #region Performance Control Methods
        /// <summary>
        /// Graphics update control is not available on ModelDocExtension in SolidWorks 2022.
        /// This method is a placeholder for future versions or alternative implementations.
        /// Use CommandInProgress and FeatureTree controls instead.
        /// </summary>
        public static void SetGraphicsUpdate(IModelDoc2 model, bool enabled)
        {
            // Note: EnableGraphicsUpdate property is not available on ModelDocExtension in SW 2022 API
            // This method is kept for API compatibility but has no effect
            ErrorHandler.DebugLog($"[PERF] SetGraphicsUpdate({enabled}) - no-op in SW 2022");
        }

        /// <summary>
        /// Sets CommandInProgress flag on the application.
        /// Setting true prevents undo record consolidation during batch ops.
        /// IMPORTANT: Always set to false when done.
        /// </summary>
        public static void SetCommandInProgress(ISldWorks swApp, bool inProgress)
        {
            const string procName = "SetCommandInProgress";
            try
            {
                if (swApp == null) return;
                swApp.CommandInProgress = inProgress;
                ErrorHandler.DebugLog($"[PERF] CommandInProgress set to {inProgress}");
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Failed to set CommandInProgress={inProgress}", ex, ErrorHandler.LogLevel.Warning);
            }
        }

        /// <summary>
        /// Enables or disables feature tree updates.
        /// Disable during feature-heavy operations for speedup.
        /// IMPORTANT: Always restore to true when done.
        /// </summary>
        public static void SetFeatureTreeUpdate(IModelDoc2 model, bool enabled)
        {
            const string procName = "SetFeatureTreeUpdate";
            try
            {
                if (model == null) return;
                var fm = model.FeatureManager;
                if (fm != null)
                {
                    fm.EnableFeatureTree = enabled;
                    ErrorHandler.DebugLog($"[PERF] FeatureTree updates set to {enabled}");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Failed to set EnableFeatureTree={enabled}", ex, ErrorHandler.LogLevel.Warning);
            }
        }
        #endregion

        // Mass Properties Methods moved to SwMassPropertiesHelper.cs
        // Sketch Methods moved to SwSketchHelper.cs
        // Custom Properties Methods moved to SwPropertyHelper.cs
        // Document IO Methods moved to SwDocumentHelper.cs
        // Custom Property Helpers moved to SwPropertyHelper.cs
        // GetFixedFace moved to SwGeometryHelper.cs

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
        #endregion
    }
}
