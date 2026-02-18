using System;
using System.Collections.Generic;
using System.IO;
using NM.Core;
using NM.Core.Config;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using static NM.Core.Constants.UnitConversions;

namespace NM.SwAddin.Drawing
{
    /// <summary>
    /// Generates drawings for parts, including flat pattern views for sheet metal.
    /// Ported from VBA SP.bas CreateDrawing() and SingleDrawing() functions.
    /// </summary>
    public sealed class DrawingGenerator
    {
        private readonly ISldWorks _swApp;

        // Layout constants (meters) from VBA SP.bas
        private const double LeftMargin = 0.0445;    // ~1.75"
        private const double BottomMargin = 0.0603;  // ~2.37"
        private const double MaxRight = 0.268;       // ~10.55" — right edge limit
        private const double MaxTop = 0.21;          // ~8.27" — top edge limit
        private const double ViewGap = 0.00635;      // ~0.25" gap between views
        private const double MetersToIn = 39.3700787401575;

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
            /// Whether to add formed (3D) isometric view for sheet metal parts.
            /// </summary>
            public bool IncludeFormedView { get; set; } = true;

            /// <summary>
            /// Whether to add a side/length view for tube parts.
            /// </summary>
            public bool IncludeSideView { get; set; } = true;

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
            public int ViewCount { get; set; }
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
                    // Use configured template path, fall back to SW default
                    templatePath = NmConfigProvider.Current?.Paths?.DrawingTemplatePath;
                    if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
                        templatePath = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
                }

                var drawDoc = _swApp.NewDocument(templatePath, (int)swDwgPaperSizes_e.swDwgPaperAsize, 0.2159, 0.2794) as IDrawingDoc;
                if (drawDoc == null)
                {
                    result.Message = "Failed to create drawing document";
                    return result;
                }

                var drawModel = drawDoc as IModelDoc2;

                // Generate view palette (populates the palette sidebar)
                bool viewsGenerated = drawDoc.GenerateViewPaletteViews(modelPath);
                if (!viewsGenerated)
                {
                    result.Message = "Failed to generate view palette";
                    return result;
                }

                // Drop the primary view at origin (0,0) — VBA drops at (0,0,0)
                IView droppedView = null;
                string droppedFromPalette = null;
                if (isSheetMetal && options.IncludeFlatPattern)
                {
                    droppedView = drawDoc.DropDrawingViewFromPalette2("Flat Pattern", 0, 0, 0) as IView;
                    if (droppedView != null) droppedFromPalette = "Flat Pattern";
                    if (droppedView == null)
                    {
                        droppedView = drawDoc.DropDrawingViewFromPalette2("*Flat pattern", 0, 0, 0) as IView;
                        if (droppedView != null) droppedFromPalette = "*Flat pattern";
                    }
                }

                if (droppedView == null)
                {
                    droppedView = drawDoc.DropDrawingViewFromPalette2("*Right", 0, 0, 0) as IView;
                    if (droppedView != null) droppedFromPalette = "*Right";
                    if (droppedView == null)
                    {
                        droppedView = drawDoc.DropDrawingViewFromPalette2("*Front", 0, 0, 0) as IView;
                        if (droppedView != null) droppedFromPalette = "*Front";
                    }
                }

                if (droppedView == null)
                {
                    result.Message = "Failed to create drawing view";
                    return result;
                }

                string droppedViewName = droppedView.GetName2();

                // Log what we dropped and where it ended up
                var dropPos = droppedView.Position as double[];
                var dropOutline = droppedView.GetOutline() as double[];
                ErrorHandler.DebugLog($"[DWG] Dropped '{droppedFromPalette}' as '{droppedViewName}', " +
                    $"center=({(dropPos != null ? dropPos[0] * MetersToIn : 0):F3}\", {(dropPos != null ? dropPos[1] * MetersToIn : 0):F3}\"), " +
                    $"outline=({(dropOutline != null ? dropOutline[0] * MetersToIn : 0):F3}\", {(dropOutline != null ? dropOutline[1] * MetersToIn : 0):F3}\") to " +
                    $"({(dropOutline != null ? dropOutline[2] * MetersToIn : 0):F3}\", {(dropOutline != null ? dropOutline[3] * MetersToIn : 0):F3}\")");

                // --- Phase 2: Drop secondary view from palette BEFORE deleting extras ---
                // Must drop from palette while it's still populated; DeleteAllViewsExcept
                // may consume or invalidate the palette state.
                IView secondaryView = null;
                int viewCount = 1;

                if (isSheetMetal && options.IncludeFormedView)
                {
                    // Drop isometric view showing formed/3D state
                    secondaryView = drawDoc.DropDrawingViewFromPalette2("*Isometric", 0.2, 0.15, 0) as IView;
                    if (secondaryView == null)
                        secondaryView = drawDoc.DropDrawingViewFromPalette2("*Front", 0.2, 0.15, 0) as IView;

                    if (secondaryView != null)
                    {
                        // Set to Default config (formed state, not flat pattern)
                        secondaryView.ReferencedConfiguration = "Default";
                        secondaryView.SetDisplayMode3(false, 2, false, true); // swHIDDEN = 2
                        viewCount++;
                        ErrorHandler.DebugLog($"[DWG] Dropped isometric as '{secondaryView.GetName2()}'");
                    }
                    else
                    {
                        ErrorHandler.DebugLog("[DWG] Failed to drop *Isometric or *Front from palette");
                    }
                }
                else if (!isSheetMetal && options.IncludeSideView)
                {
                    // Tube: project side view from primary end view
                    secondaryView = AddTubeSideView(drawDoc, drawModel, droppedView);
                    if (secondaryView != null)
                        viewCount++;
                }

                // Track all view names to keep
                var keepViews = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { droppedViewName };
                if (secondaryView != null)
                    keepViews.Add(secondaryView.GetName2());

                // Delete all auto-inserted views EXCEPT the ones we want to keep.
                // GenerateViewPaletteViews auto-populates the sheet with standard views.
                DeleteAllViewsExcept(drawDoc, drawModel, keepViews);

                // Rebuild to get accurate outline before positioning (VBA does this)
                drawModel.ForceRebuild3(false);

                // Position primary view on sheet (match VBA: left margin + clamp to sheet bounds)
                PositionView(drawDoc, droppedView);
                drawModel.EditRebuild3();

                // Check grain direction (for sheet metal)
                if (isSheetMetal)
                {
                    CheckAndSetGrainDirection(model, droppedView);
                }

                // Rotate primary view if needed (portrait orientation)
                RotateViewIfNeeded(drawDoc, droppedView);

                // Position secondary view relative to primary
                if (secondaryView != null)
                {
                    if (isSheetMetal)
                        PositionSecondaryViewUpperRight(drawDoc, secondaryView, droppedView);
                    // Tube side view is already positioned by AddTubeSideView
                }

                drawModel.EditRebuild3();
                result.ViewCount = viewCount;

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
                var drawModel = drawDoc as IModelDoc2;

                // Step 1: Position view so left edge is at LeftMargin, bottom at BottomMargin
                // Use delta-based positioning: shift the view center by how far the edge
                // needs to move to reach the margin. This works regardless of where
                // SolidWorks initially placed the view after dropping.
                var outline = view.GetOutline() as double[];
                if (outline == null || outline.Length < 4) return;

                var position = view.Position as double[];
                if (position == null || position.Length < 2) return;

                ErrorHandler.DebugLog($"[DWG] PositionView BEFORE: center=({position[0] * MetersToIn:F3}\", {position[1] * MetersToIn:F3}\") " +
                    $"outline=({outline[0] * MetersToIn:F3}\", {outline[1] * MetersToIn:F3}\") to ({outline[2] * MetersToIn:F3}\", {outline[3] * MetersToIn:F3}\")");

                // Delta-based: move center by the difference between desired edge and current edge
                double deltaX = LeftMargin - outline[0];
                double deltaY = BottomMargin - outline[1];
                position[0] += deltaX;
                position[1] += deltaY;
                view.Position = position;

                // VBA calls EditRebuild after initial positioning to update the outline
                if (drawModel != null)
                    drawModel.EditRebuild3();

                // Step 2: Clamp right edge — if view extends past MaxRight, shift left
                outline = view.GetOutline() as double[];
                if (outline != null && outline[2] > MaxRight)
                {
                    position = view.Position as double[];
                    if (position != null)
                    {
                        position[0] -= (outline[2] - MaxRight);
                        view.Position = position;
                    }
                }

                // Step 3: Clamp top edge — if view extends past MaxTop, shift down
                outline = view.GetOutline() as double[];
                if (outline != null && outline[3] > MaxTop)
                {
                    position = view.Position as double[];
                    if (position != null)
                    {
                        position[1] -= (outline[3] - MaxTop);
                        view.Position = position;
                    }
                }

                // Log final position
                outline = view.GetOutline() as double[];
                position = view.Position as double[];
                if (outline != null && position != null)
                {
                    ErrorHandler.DebugLog($"[DWG] PositionView AFTER: center=({position[0] * MetersToIn:F3}\", {position[1] * MetersToIn:F3}\") " +
                        $"outline=({outline[0] * MetersToIn:F3}\", {outline[1] * MetersToIn:F3}\") to ({outline[2] * MetersToIn:F3}\", {outline[3] * MetersToIn:F3}\")");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"PositionView: {ex.Message}");
            }
        }

        /// <summary>
        /// Positions a secondary view in the upper-right area of the sheet,
        /// avoiding overlap with the primary view and staying within sheet bounds.
        /// </summary>
        private void PositionSecondaryViewUpperRight(IDrawingDoc drawDoc, IView secondaryView, IView primaryView)
        {
            try
            {
                var drawModel = drawDoc as IModelDoc2;

                var primaryOutline = primaryView.GetOutline() as double[];
                var secOutline = secondaryView.GetOutline() as double[];
                if (secOutline == null || secOutline.Length < 4) return;

                var secPos = secondaryView.Position as double[];
                if (secPos == null || secPos.Length < 2) return;

                double secWidth = secOutline[2] - secOutline[0];
                double secHeight = secOutline[3] - secOutline[1];

                // Target: right edge at MaxRight, top edge at MaxTop
                double targetCenterX = MaxRight - (secWidth / 2.0);
                double targetCenterY = MaxTop - (secHeight / 2.0);

                // Delta-based positioning
                secPos[0] += (targetCenterX - secPos[0]);
                secPos[1] += (targetCenterY - secPos[1]);
                secondaryView.Position = secPos;

                if (drawModel != null) drawModel.EditRebuild3();

                // Check for overlap with primary view and shrink if needed
                if (primaryOutline != null && primaryOutline.Length >= 4)
                {
                    for (int attempt = 0; attempt < 5; attempt++)
                    {
                        secOutline = secondaryView.GetOutline() as double[];
                        if (secOutline == null) break;

                        // Check if secondary left edge overlaps primary right edge
                        bool overlapsX = secOutline[0] < primaryOutline[2] + ViewGap;
                        // Check if secondary bottom edge overlaps primary top edge
                        bool overlapsY = secOutline[1] < primaryOutline[3] + ViewGap;

                        if (!overlapsX || !overlapsY) break; // No overlap in both dimensions needed

                        // Shrink the view scale by 15%
                        double[] scaleRatio = secondaryView.ScaleRatio as double[];
                        if (scaleRatio != null && scaleRatio.Length >= 2 && scaleRatio[1] > 0)
                        {
                            double currentScale = scaleRatio[0] / scaleRatio[1];
                            double newNum = scaleRatio[0] * 0.85;
                            secondaryView.ScaleRatio = new double[] { newNum, scaleRatio[1] };
                        }

                        if (drawModel != null) drawModel.EditRebuild3();

                        // Reposition after scale change
                        secOutline = secondaryView.GetOutline() as double[];
                        if (secOutline == null) break;
                        secWidth = secOutline[2] - secOutline[0];
                        secHeight = secOutline[3] - secOutline[1];
                        targetCenterX = MaxRight - (secWidth / 2.0);
                        targetCenterY = MaxTop - (secHeight / 2.0);
                        secPos = secondaryView.Position as double[];
                        if (secPos == null) break;
                        secPos[0] = targetCenterX;
                        secPos[1] = targetCenterY;
                        secondaryView.Position = secPos;
                        if (drawModel != null) drawModel.EditRebuild3();
                    }
                }

                // Final clamp to sheet bounds
                secOutline = secondaryView.GetOutline() as double[];
                secPos = secondaryView.Position as double[];
                if (secOutline != null && secPos != null)
                {
                    if (secOutline[2] > MaxRight)
                        secPos[0] -= (secOutline[2] - MaxRight);
                    if (secOutline[3] > MaxTop)
                        secPos[1] -= (secOutline[3] - MaxTop);
                    if (secOutline[0] < LeftMargin)
                        secPos[0] += (LeftMargin - secOutline[0]);
                    if (secOutline[1] < BottomMargin)
                        secPos[1] += (BottomMargin - secOutline[1]);
                    secondaryView.Position = secPos;
                }

                secOutline = secondaryView.GetOutline() as double[];
                secPos = secondaryView.Position as double[];
                if (secOutline != null && secPos != null)
                {
                    ErrorHandler.DebugLog($"[DWG] IsoView FINAL: center=({secPos[0] * MetersToIn:F3}\",{secPos[1] * MetersToIn:F3}\") " +
                        $"outline=({secOutline[0] * MetersToIn:F3}\",{secOutline[1] * MetersToIn:F3}\") to " +
                        $"({secOutline[2] * MetersToIn:F3}\",{secOutline[3] * MetersToIn:F3}\")");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"PositionSecondaryViewUpperRight: {ex.Message}");
            }
        }

        /// <summary>
        /// Adds a side profile view for tube parts showing the tube length.
        /// Uses CreateUnfoldedViewAt3 to project from the primary end view.
        /// Ported from VBA SP.bas lines 1811-1828.
        /// </summary>
        private IView AddTubeSideView(IDrawingDoc drawDoc, IModelDoc2 drawModel, IView primaryView)
        {
            const string proc = nameof(AddTubeSideView);
            try
            {
                string primaryName = primaryView.GetName2();

                // Activate and select the primary view (required for CreateUnfoldedViewAt3)
                drawDoc.ActivateView(primaryName);
                drawModel.Extension.SelectByID2(primaryName, "DRAWINGVIEW",
                    0, 0, 0, false, 0, null, 0);

                // Get primary view position for Y alignment
                var primaryPos = primaryView.Position as double[];
                if (primaryPos == null || primaryPos.Length < 2) return null;

                // Create projected view to the right (VBA: 0.3 meters from left, same Y)
                var sideView = drawDoc.CreateUnfoldedViewAt3(0.2, primaryPos[1], 0, false) as IView;
                if (sideView == null)
                {
                    ErrorHandler.DebugLog($"{proc}: CreateUnfoldedViewAt3 returned null");
                    return null;
                }

                drawModel.ClearSelection2(true);

                // Set to Default config and hidden lines visible
                sideView.ReferencedConfiguration = "Default";
                sideView.SetDisplayMode3(false, 2, false, true);

                drawModel.ForceRebuild3(false);

                // Position side view to the right of primary
                PositionTubeSideView(drawDoc, primaryView, sideView);

                ErrorHandler.DebugLog($"{proc}: Added '{sideView.GetName2()}'");
                return sideView;
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"{proc}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Positions the tube side view to the right of the primary end view.
        /// Primary stays left-aligned, side view goes to the right with a gap.
        /// </summary>
        private void PositionTubeSideView(IDrawingDoc drawDoc, IView primaryView, IView sideView)
        {
            try
            {
                var drawModel = drawDoc as IModelDoc2;

                // Get primary outline after positioning
                var primaryOutline = primaryView.GetOutline() as double[];
                var sideOutline = sideView.GetOutline() as double[];
                var sidePos = sideView.Position as double[];

                if (primaryOutline == null || sideOutline == null || sidePos == null) return;

                // Position side view: left edge at primary right edge + gap, Y centered with primary
                var primaryPos = primaryView.Position as double[];
                double targetLeft = primaryOutline[2] + ViewGap;
                double sideLeftEdge = sideOutline[0];
                sidePos[0] += (targetLeft - sideLeftEdge);
                if (primaryPos != null)
                    sidePos[1] = primaryPos[1];
                sideView.Position = sidePos;

                if (drawModel != null) drawModel.EditRebuild3();

                // Clamp right edge
                sideOutline = sideView.GetOutline() as double[];
                if (sideOutline != null && sideOutline[2] > MaxRight)
                {
                    sidePos = sideView.Position as double[];
                    if (sidePos != null)
                    {
                        sidePos[0] -= (sideOutline[2] - MaxRight);
                        sideView.Position = sidePos;
                    }
                }

                // Clamp top/bottom
                sideOutline = sideView.GetOutline() as double[];
                if (sideOutline != null)
                {
                    sidePos = sideView.Position as double[];
                    if (sidePos != null)
                    {
                        if (sideOutline[3] > MaxTop)
                            sidePos[1] -= (sideOutline[3] - MaxTop);
                        if (sideOutline[1] < BottomMargin)
                            sidePos[1] += (BottomMargin - sideOutline[1]);
                        sideView.Position = sidePos;
                    }
                }

                // Log positions
                primaryOutline = primaryView.GetOutline() as double[];
                sideOutline = sideView.GetOutline() as double[];
                ErrorHandler.DebugLog($"[DWG] TubeLayout: primary=({(primaryOutline?[0] ?? 0) * MetersToIn:F2}\"-{(primaryOutline?[2] ?? 0) * MetersToIn:F2}\"), " +
                    $"side=({(sideOutline?[0] ?? 0) * MetersToIn:F2}\"-{(sideOutline?[2] ?? 0) * MetersToIn:F2}\")");
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"PositionTubeSideView: {ex.Message}");
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
                double maxXInch = maxX * MetersToIn;
                double maxYInch = maxY * MetersToIn;

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

        /// <summary>
        /// Deletes all views from the current sheet except the named ones.
        /// GenerateViewPaletteViews may auto-populate the sheet with standard views on some templates.
        /// </summary>
        private void DeleteAllViewsExcept(IDrawingDoc drawDoc, IModelDoc2 drawModel, HashSet<string> keepViewNames)
        {
            try
            {
                var sheet = drawDoc.GetCurrentSheet() as ISheet;
                if (sheet == null) return;

                var viewsRaw = sheet.GetViews() as object[];
                if (viewsRaw == null || viewsRaw.Length <= 1) return;

                var toDelete = new List<string>();
                foreach (var viewObj in viewsRaw)
                {
                    var view = viewObj as IView;
                    if (view == null) continue;

                    string viewName = view.GetName2();
                    if (string.IsNullOrEmpty(viewName)) continue;
                    if (keepViewNames.Contains(viewName)) continue;

                    toDelete.Add(viewName);
                }

                foreach (string viewName in toDelete)
                {
                    drawModel.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);
                    drawModel.DeleteSelection(false);
                }

                drawModel.ClearSelection2(true);
                ErrorHandler.DebugLog($"DeleteAllViewsExcept: deleted {toDelete.Count} views, kept {keepViewNames.Count}");
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"DeleteAllViewsExcept: {ex.Message}");
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
