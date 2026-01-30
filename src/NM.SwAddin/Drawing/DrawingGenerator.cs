using System;
using System.IO;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Drawing
{
    /// <summary>
    /// Generates drawings for parts, including flat pattern views for sheet metal.
    /// Ported from VBA SP.bas CreateDrawing() and SingleDrawing() functions.
    /// </summary>
    public sealed class DrawingGenerator
    {
        private readonly ISldWorks _swApp;

        /// <summary>
        /// Drawing generation options.
        /// </summary>
        public sealed class DrawingOptions
        {
            /// <summary>
            /// Path to drawing template (.drwdot file).
            /// If null/empty, uses SolidWorks default.
            /// </summary>
            public string TemplatePath { get; set; }

            /// <summary>
            /// Whether to create DXF output alongside the drawing.
            /// </summary>
            public bool CreateDxf { get; set; } = true;

            /// <summary>
            /// Whether to save the drawing file.
            /// </summary>
            public bool SaveDrawing { get; set; } = true;

            /// <summary>
            /// Whether to add flat pattern view for sheet metal parts.
            /// </summary>
            public bool IncludeFlatPattern { get; set; } = true;

            /// <summary>
            /// Whether to add formed (3D) view.
            /// </summary>
            public bool IncludeFormedView { get; set; } = true;

            /// <summary>
            /// Custom output folder. If null, uses same folder as part.
            /// </summary>
            public string OutputFolder { get; set; }
        }

        /// <summary>
        /// Result of drawing generation.
        /// </summary>
        public sealed class DrawingResult
        {
            public bool Success { get; set; }
            public string DrawingPath { get; set; }
            public string DxfPath { get; set; }
            public string Message { get; set; }
            public bool WasExisting { get; set; }
        }

        public DrawingGenerator(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Creates or opens a drawing for the active document.
        /// Equivalent to VBA SingleDrawing() function.
        /// </summary>
        public DrawingResult CreateDrawingForActiveDoc(DrawingOptions options = null)
        {
            const string proc = nameof(CreateDrawingForActiveDoc);
            ErrorHandler.PushCallStack(proc);
            options = options ?? new DrawingOptions();

            try
            {
                var model = _swApp.ActiveDoc as IModelDoc2;
                if (model == null)
                {
                    return new DrawingResult { Success = false, Message = "No active document" };
                }

                return CreateDrawing(model, options);
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                return new DrawingResult { Success = false, Message = "Exception: " + ex.Message };
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Creates or opens a drawing for the specified model.
        /// Equivalent to VBA CreateDrawing() function.
        /// </summary>
        public DrawingResult CreateDrawing(IModelDoc2 model, DrawingOptions options = null)
        {
            const string proc = nameof(CreateDrawing);
            ErrorHandler.PushCallStack(proc);
            options = options ?? new DrawingOptions();

            var result = new DrawingResult();

            try
            {
                if (model == null)
                {
                    result.Message = "Model is null";
                    return result;
                }

                var docType = (swDocumentTypes_e)model.GetType();
                if (docType != swDocumentTypes_e.swDocPART)
                {
                    result.Message = "Document must be a part";
                    return result;
                }

                // Get paths
                string modelPath = model.GetPathName();
                string modelTitle = model.GetTitle() ?? "Untitled";
                string docName = Path.GetFileNameWithoutExtension(modelTitle);
                string folderPath = !string.IsNullOrEmpty(modelPath)
                    ? Path.GetDirectoryName(modelPath)
                    : System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);

                if (!string.IsNullOrEmpty(options.OutputFolder))
                    folderPath = options.OutputFolder;

                string drawingPath = Path.Combine(folderPath, docName + ".slddrw");
                string dxfPath = Path.Combine(folderPath, docName + ".dxf");

                // Check if drawing already exists
                if (File.Exists(drawingPath))
                {
                    int errors = 0;
                    int warnings = 0;
                    _swApp.OpenDoc6(drawingPath, (int)swDocumentTypes_e.swDocDRAWING, 0, "", ref errors, ref warnings);
                    result.Success = true;
                    result.WasExisting = true;
                    result.DrawingPath = drawingPath;
                    result.Message = "Existing drawing opened";
                    return result;
                }

                // Check if part has been processed (has OP20 property)
                string op20 = GetCustomProperty(model, "OP20");
                if (string.IsNullOrEmpty(op20))
                {
                    result.Message = "Part has not been processed (no OP20 property)";
                    return result;
                }

                // Save the model first
                int saveErrors = 0;
                int saveWarnings = 0;
                model.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref saveErrors, ref saveWarnings);

                // Check if this is a sheet metal or laser part
                bool isSheetMetal = op20.StartsWith("N115") || op20.StartsWith("N120") || op20.StartsWith("N125") ||
                                    op20.StartsWith("F115") || op20.StartsWith("F110");
                bool isTubeLaser = op20.Contains("TUBE LASER") || op20.StartsWith("F110");

                // Prepare flat pattern configuration if needed
                if (isSheetMetal && options.IncludeFlatPattern)
                {
                    PrepareFlatPatternConfig(model);
                }

                // Create new drawing
                string templatePath = options.TemplatePath;
                if (string.IsNullOrEmpty(templatePath))
                {
                    templatePath = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
                }

                var drawDoc = _swApp.NewDocument(templatePath, (int)swDwgPaperSizes_e.swDwgPaperAsize, 0.2159, 0.2794) as IDrawingDoc;
                if (drawDoc == null)
                {
                    result.Message = "Failed to create drawing document";
                    return result;
                }

                var drawModel = drawDoc as IModelDoc2;

                // Generate view palette
                bool viewsGenerated = drawDoc.GenerateViewPaletteViews(modelPath);
                if (!viewsGenerated)
                {
                    result.Message = "Failed to generate view palette";
                    return result;
                }

                // Drop the primary view
                IView droppedView = null;
                if (isSheetMetal && options.IncludeFlatPattern)
                {
                    // Try to drop flat pattern view first
                    droppedView = drawDoc.DropDrawingViewFromPalette2("Flat Pattern", 0.1, 0.1, 0) as IView;
                    if (droppedView == null)
                    {
                        droppedView = drawDoc.DropDrawingViewFromPalette2("*Flat pattern", 0.1, 0.1, 0) as IView;
                    }
                }

                if (droppedView == null)
                {
                    // Fall back to Right view or first available
                    droppedView = drawDoc.DropDrawingViewFromPalette2("*Right", 0.1, 0.1, 0) as IView;
                    if (droppedView == null)
                    {
                        droppedView = drawDoc.DropDrawingViewFromPalette2("*Front", 0.1, 0.1, 0) as IView;
                    }
                }

                if (droppedView == null)
                {
                    result.Message = "Failed to create drawing view";
                    return result;
                }

                // Rebuild
                drawModel.ForceRebuild3(false);

                // Position and scale the view
                PositionView(drawDoc, droppedView);

                // Check grain direction (for sheet metal)
                if (isSheetMetal)
                {
                    CheckAndSetGrainDirection(model, droppedView);
                }

                // Rotate view if needed (portrait orientation)
                RotateViewIfNeeded(drawDoc, droppedView);

                // Save DXF if requested
                if (options.CreateDxf)
                {
                    int dxfErrors = 0;
                    int dxfWarnings = 0;
                    bool dxfSaved = drawModel.SaveAs4(
                        dxfPath,
                        (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                        ref dxfErrors,
                        ref dxfWarnings);
                    if (dxfSaved)
                        result.DxfPath = dxfPath;
                }

                // Save drawing if requested
                if (options.SaveDrawing)
                {
                    int drwErrors = 0;
                    int drwWarnings = 0;
                    bool drwSaved = drawModel.SaveAs4(
                        drawingPath,
                        (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                        ref drwErrors,
                        ref drwWarnings);
                    if (drwSaved)
                        result.DrawingPath = drawingPath;
                }

                // Save the part (in case grain property was updated)
                model.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref saveErrors, ref saveWarnings);

                result.Success = true;
                result.Message = "Drawing created successfully";
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

        private void PrepareFlatPatternConfig(IModelDoc2 model)
        {
            try
            {
                var confMgr = model.ConfigurationManager;
                if (confMgr == null) return;

                // Check if flat pattern config already exists
                var configNames = model.GetConfigurationNames() as string[];
                bool hasFlatConfig = false;
                if (configNames != null)
                {
                    foreach (var name in configNames)
                    {
                        if (name != null && name.ToUpperInvariant().Contains("FLAT"))
                        {
                            hasFlatConfig = true;
                            break;
                        }
                    }
                }

                if (hasFlatConfig) return;

                // Find and unsuppress flat pattern feature
                var feat = model.FirstFeature() as IFeature;
                while (feat != null)
                {
                    string typeName = feat.GetTypeName2() ?? "";
                    if (typeName.Equals("FlatPattern", StringComparison.OrdinalIgnoreCase))
                    {
                        feat.SetSuppression2(
                            (int)swFeatureSuppressionAction_e.swUnSuppressFeature,
                            (int)swInConfigurationOpts_e.swThisConfiguration,
                            null);

                        // Create flat pattern configuration
                        confMgr.AddConfiguration(
                            "DefaultSM-FLAT-PATTERN",
                            "Flattened state of sheet metal part",
                            "",
                            0,
                            "Default",
                            "DefaultSM-FLAT-PATTERN");

                        // Switch back to default and suppress flat pattern
                        model.ShowConfiguration2("Default");
                        feat.SetSuppression2(
                            (int)swFeatureSuppressionAction_e.swSuppressFeature,
                            (int)swInConfigurationOpts_e.swThisConfiguration,
                            null);

                        break;
                    }
                    feat = feat.GetNextFeature() as IFeature;
                }

                model.EditRebuild3();
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"PrepareFlatPatternConfig: {ex.Message}");
            }
        }

        private void PositionView(IDrawingDoc drawDoc, IView view)
        {
            try
            {
                var outline = view.GetOutline() as double[];
                if (outline == null || outline.Length < 4) return;

                // Get sheet size
                var sheet = drawDoc.GetCurrentSheet() as ISheet;
                if (sheet == null) return;

                var props = sheet.GetProperties2() as double[];
                if (props == null) return;

                // Position view near bottom-left with margin
                double margin = 0.0445; // ~1.75"
                var position = view.Position as double[];
                if (position != null && position.Length >= 2)
                {
                    position[0] = margin - outline[0];
                    position[1] = 0.0603 - outline[1];
                    view.Position = position;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"PositionView: {ex.Message}");
            }
        }

        private void CheckAndSetGrainDirection(IModelDoc2 model, IView view)
        {
            try
            {
                var outline = view.GetOutline() as double[];
                if (outline == null || outline.Length < 4) return;

                double maxX = outline[2] - outline[0];
                double maxY = outline[3] - outline[1];

                // Convert to inches (from meters)
                double maxXInch = maxX * 39.3701;
                double maxYInch = maxY * 39.3701;

                // If either dimension is under 6", mark as grain-sensitive
                if (maxXInch <= 6 || maxYInch <= 6)
                {
                    string configName = model.ConfigurationManager?.ActiveConfiguration?.Name ?? "";
                    SetCustomProperty(model, configName, "Grain", "Y");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"CheckAndSetGrainDirection: {ex.Message}");
            }
        }

        private void RotateViewIfNeeded(IDrawingDoc drawDoc, IView view)
        {
            try
            {
                var outline = view.GetOutline() as double[];
                if (outline == null || outline.Length < 4) return;

                double maxX = outline[2] - outline[0];
                double maxY = outline[3] - outline[1];

                // If taller than wide, rotate 90 degrees
                if (maxY > maxX)
                {
                    var drawModel = drawDoc as IModelDoc2;
                    if (drawModel == null) return;

                    drawDoc.ActivateView("Drawing View1");
                    drawModel.Extension.SelectByID2("Drawing View1", "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                    drawDoc.DrawingViewRotate(Math.PI / 2); // 90 degrees
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"RotateViewIfNeeded: {ex.Message}");
            }
        }

        private string GetCustomProperty(IModelDoc2 model, string propName)
        {
            try
            {
                string configName = model.ConfigurationManager?.ActiveConfiguration?.Name ?? "";
                var propMgr = model.Extension?.CustomPropertyManager[configName];
                if (propMgr == null)
                    propMgr = model.Extension?.CustomPropertyManager[""];

                if (propMgr != null)
                {
                    string valOut = "";
                    string resolvedOut = "";
                    bool wasResolved = false;
                    propMgr.Get5(propName, true, out valOut, out resolvedOut, out wasResolved);
                    return resolvedOut ?? valOut ?? "";
                }
            }
            catch { }
            return "";
        }

        private void SetCustomProperty(IModelDoc2 model, string configName, string propName, string value)
        {
            try
            {
                var propMgr = model.Extension?.CustomPropertyManager[configName];
                if (propMgr != null)
                {
                    propMgr.Add3(propName, (int)swCustomInfoType_e.swCustomInfoText, value, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                }
            }
            catch { }
        }

        #region Flat Pattern DXF Export for Nesting

        /// <summary>
        /// Options for flat pattern DXF export.
        /// </summary>
        public sealed class FlatPatternDxfOptions
        {
            /// <summary>
            /// Output folder. If null, uses same folder as part.
            /// </summary>
            public string OutputFolder { get; set; }

            /// <summary>
            /// Whether to include bend lines in the DXF.
            /// Default false - nesting software typically only needs cut geometry.
            /// </summary>
            public bool IncludeBendLines { get; set; } = false;

            /// <summary>
            /// Whether to include sketches in the DXF.
            /// </summary>
            public bool IncludeSketches { get; set; } = false;

            /// <summary>
            /// Whether to include library features (forming tools).
            /// </summary>
            public bool IncludeLibraryFeatures { get; set; } = false;

            /// <summary>
            /// Whether to include forming tool extents.
            /// </summary>
            public bool IncludeFormingToolExtents { get; set; } = false;

            /// <summary>
            /// DXF version (12-14 for maximum compatibility with nesting software).
            /// </summary>
            public int DxfVersion { get; set; } = 13; // R13 - widely compatible
        }

        /// <summary>
        /// Result of flat pattern DXF export.
        /// </summary>
        public sealed class FlatPatternDxfResult
        {
            public bool Success { get; set; }
            public string DxfPath { get; set; }
            public string Message { get; set; }
        }

        /// <summary>
        /// Exports the flat pattern of a sheet metal part directly to DXF for nesting software.
        /// Uses IPartDoc.ExportFlatPatternView() which creates 2D geometry suitable for CNC nesting.
        /// </summary>
        /// <param name="model">The sheet metal part model.</param>
        /// <param name="options">Export options (optional).</param>
        /// <returns>Result with path to the exported DXF.</returns>
        public FlatPatternDxfResult ExportFlatPatternDxf(IModelDoc2 model, FlatPatternDxfOptions options = null)
        {
            const string proc = nameof(ExportFlatPatternDxf);
            ErrorHandler.PushCallStack(proc);
            options = options ?? new FlatPatternDxfOptions();
            var result = new FlatPatternDxfResult();

            try
            {
                if (model == null)
                {
                    result.Message = "Model is null";
                    return result;
                }

                var partDoc = model as IPartDoc;
                if (partDoc == null)
                {
                    result.Message = "Document must be a part";
                    return result;
                }

                // Verify this is a sheet metal part with flat pattern
                if (!HasFlatPatternFeature(model))
                {
                    result.Message = "Part does not have a flat pattern feature (not sheet metal)";
                    return result;
                }

                // Get output path
                string modelPath = model.GetPathName();
                string modelTitle = model.GetTitle() ?? "Untitled";
                string docName = Path.GetFileNameWithoutExtension(
                    !string.IsNullOrEmpty(modelPath) ? modelPath : modelTitle);

                string folderPath = !string.IsNullOrEmpty(modelPath)
                    ? Path.GetDirectoryName(modelPath)
                    : System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);

                if (!string.IsNullOrEmpty(options.OutputFolder))
                    folderPath = options.OutputFolder;

                string dxfPath = Path.Combine(folderPath, docName + "_FLAT.dxf");

                // Build export options bitmask
                // swExportFlatPatternViewOptions_e flags:
                // swExportFlatPatternOption_RemoveBends = 1 (remove bend lines)
                // swExportFlatPatternOption_ExportFlatPatternGeometry = 2 (geometry only, no annotations)
                // swExportFlatPatternOption_IncludeHiddenEdges = 4
                // swExportFlatPatternOption_ExportBendLines = 8
                // swExportFlatPatternOption_ExportSketches = 16
                // swExportFlatPatternOption_ExportLibraryFeatures = 32
                // swExportFlatPatternOption_ExportFormingToolExtents = 64
                // swExportFlatPatternOption_ExportBendNotes = 128

                int exportOptions = 2; // Always include flat pattern geometry

                if (options.IncludeBendLines)
                    exportOptions |= 8;
                else
                    exportOptions |= 1; // Remove bends if not including them

                if (options.IncludeSketches)
                    exportOptions |= 16;

                if (options.IncludeLibraryFeatures)
                    exportOptions |= 32;

                if (options.IncludeFormingToolExtents)
                    exportOptions |= 64;

                ErrorHandler.DebugLog($"{proc}: Exporting flat pattern to {dxfPath} with options={exportOptions}");

                // Ensure flat pattern is unsuppressed
                UnsuppressFlatPattern(model);

                // Export the flat pattern view
                bool success = partDoc.ExportFlatPatternView(dxfPath, exportOptions);

                if (success && File.Exists(dxfPath))
                {
                    result.Success = true;
                    result.DxfPath = dxfPath;
                    result.Message = "Flat pattern DXF exported successfully";
                    ErrorHandler.DebugLog($"{proc}: Success - {dxfPath}");
                }
                else
                {
                    result.Message = "ExportFlatPatternView returned false or file not created";
                    ErrorHandler.DebugLog($"{proc}: Export failed");
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
        /// Checks if the model has a flat pattern feature (is sheet metal).
        /// </summary>
        private bool HasFlatPatternFeature(IModelDoc2 model)
        {
            try
            {
                var feat = model.FirstFeature() as IFeature;
                while (feat != null)
                {
                    string typeName = feat.GetTypeName2() ?? "";
                    if (typeName.Equals("FlatPattern", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                    feat = feat.GetNextFeature() as IFeature;
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// Unsuppresses the flat pattern feature if suppressed.
        /// Required before ExportFlatPatternView.
        /// </summary>
        private void UnsuppressFlatPattern(IModelDoc2 model)
        {
            try
            {
                var feat = model.FirstFeature() as IFeature;
                while (feat != null)
                {
                    string typeName = feat.GetTypeName2() ?? "";
                    if (typeName.Equals("FlatPattern", StringComparison.OrdinalIgnoreCase))
                    {
                        // Always try to unsuppress - SetSuppression2 is idempotent
                        // If already unsuppressed, this is a no-op
                        feat.SetSuppression2(
                            (int)swFeatureSuppressionAction_e.swUnSuppressFeature,
                            (int)swInConfigurationOpts_e.swThisConfiguration,
                            null);
                        model.EditRebuild3();
                        break;
                    }
                    feat = feat.GetNextFeature() as IFeature;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"UnsuppressFlatPattern: {ex.Message}");
            }
        }

        #endregion
    }
}
