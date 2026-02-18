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
        private const double TubeYPos = 0.1215;      // ~4.78" — tube views vertically centered
        private const double ProjectedOffset = 0.3;  // ~11.8" — initial projected view offset
        private const double MaxRight = 0.268;       // ~10.55" — right edge limit
        private const double MaxTop = 0.21;          // ~8.27" — top edge limit
        private const double ViewGap = 0.00635;      // ~0.25" gap between views
        private const double ScaleFactor = 1.05;     // 5% scale-up per iteration
        private const int ScaleMaxIter = 25;         // Max scaling iterations
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
            /// Whether to auto-dimension drawing views.
            /// </summary>
            public bool IncludeDimensions { get; set; } = true;

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
            public int DimensionsAdded { get; set; }
            public int EtchMarksFound { get; set; }
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

                // Phase 5: Hide reference planes before creating drawing views
                int planesHidden = HideReferencePlanes(model);
                ErrorHandler.DebugLog($"[DWG] Hidden {planesHidden} reference plane(s)");

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

                // --- Phase 2: Create secondary views ---
                // Sheet metal: projected views are created AFTER dimensioning the flat pattern
                //              (VBA creates them inside DimensionFlat, not here).
                // Tube: project side view from primary end view.
                IView secondaryView = null;
                int viewCount = 1;

                if (!isSheetMetal && options.IncludeSideView)
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
                if (isTubeLaser)
                    PositionView(drawDoc, droppedView, TubeYPos);
                else
                    PositionView(drawDoc, droppedView);
                drawModel.EditRebuild3();

                // Check grain direction (for sheet metal)
                if (isSheetMetal)
                {
                    CheckAndSetGrainDirection(model, droppedView);
                }

                // Rotate primary view if needed (portrait orientation)
                RotateViewIfNeeded(drawDoc, droppedView);

                // Position tube side view (already created above)
                if (secondaryView != null && !isSheetMetal)
                {
                    // Tube side view is already positioned by AddTubeSideView
                }

                drawModel.EditRebuild3();

                // Phase 3: Auto-dimension views and create projected views for sheet metal
                if (options.IncludeDimensions)
                {
                    var dimensioner = new DrawingDimensioner(_swApp);
                    try
                    {
                        if (isSheetMetal)
                        {
                            // Dimension flat pattern (bends + overall)
                            var dimResult = dimensioner.DimensionFlatPattern(drawDoc, droppedView);
                            result.DimensionsAdded += dimResult.DimensionsAdded;

                            // Create projected views based on bend analysis (VBA does this inside DimensionFlat)
                            if (options.IncludeFormedView)
                            {
                                var projectedViews = CreateProjectedViews(
                                    drawDoc, drawModel, droppedView, dimensioner,
                                    dimResult.HasVerticalBends, dimResult.HasHorizontalBends);

                                result.DimensionsAdded += projectedViews.DimensionsAdded;
                                viewCount += projectedViews.ViewsCreated;
                            }
                        }
                        else
                        {
                            // Dimension tube primary view
                            var dimResult = dimensioner.DimensionTube(drawDoc, droppedView);
                            result.DimensionsAdded += dimResult.DimensionsAdded;

                            // Dimension tube side view with DimensionTube (not DimensionFormedView)
                            if (secondaryView != null)
                            {
                                var secDimResult = dimensioner.DimensionTube(drawDoc, secondaryView);
                                result.DimensionsAdded += secDimResult.DimensionsAdded;
                            }
                        }

                        dimensioner.AlignAllDimensions(drawDoc);
                        ErrorHandler.DebugLog($"[DWG] Dimensions added: {result.DimensionsAdded}");
                    }
                    catch (Exception dimEx)
                    {
                        ErrorHandler.DebugLog($"[DWG] Dimensioning error: {dimEx.Message}");
                    }
                }

                // Phase 5: Make etch marks visible in drawing
                int etchCount = MakeEtchMarksVisible(drawDoc, droppedView, model);
                result.EtchMarksFound = etchCount;

                result.ViewCount = viewCount;

                // Auto-scale sheet to fit all views
                AutoScaleToFit(drawDoc, drawModel);

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

        private void PositionView(IDrawingDoc drawDoc, IView view, double yPosition = BottomMargin)
        {
            try
            {
                var drawModel = drawDoc as IModelDoc2;

                // Step 1: Position view so left edge is at LeftMargin, bottom at yPosition
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
                double deltaY = yPosition - outline[1];
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

        /// <summary>
        /// Creates projected (Top/Right) views from the flat pattern view, sets them to
        /// Default config with hidden-greyed display, and dimensions them with overall W/H.
        /// Ported from VBA DimensionFlat() projected view creation (SP.bas lines 1872-1878).
        /// </summary>
        private DrawingDimensioner.DimensionResult CreateProjectedViews(
            IDrawingDoc drawDoc, IModelDoc2 drawModel, IView flatPatternView,
            DrawingDimensioner dimensioner, bool hasVerticalBends, bool hasHorizontalBends)
        {
            const string proc = nameof(CreateProjectedViews);
            var result = new DrawingDimensioner.DimensionResult();

            try
            {
                var flatPos = flatPatternView.Position as double[];
                if (flatPos == null || flatPos.Length < 2) return result;

                string flatViewName = flatPatternView.GetName2();

                // TopView: created if vertical bends exist (shows top-down formed shape)
                if (hasVerticalBends)
                {
                    try
                    {
                        drawDoc.ActivateView(flatViewName);
                        drawModel.Extension.SelectByID2(flatViewName, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

                        // VBA: CreateUnfoldedViewAt3(flatPos.X, 0.3, 0, false)
                        var topView = drawDoc.CreateUnfoldedViewAt3(flatPos[0], ProjectedOffset, 0, false) as IView;
                        if (topView != null)
                        {
                            drawModel.ClearSelection2(true);
                            topView.ReferencedConfiguration = "Default";
                            topView.SetDisplayMode3(false, 4 /* swHIDDEN_GREYED */, false, true);
                            drawModel.ForceRebuild3(false);

                            // Clamp within sheet bounds
                            ClampViewToSheet(drawDoc, topView);

                            // Dimension with overall W/H only
                            var dimResult = dimensioner.DimensionFormedView(drawDoc, topView);
                            result.DimensionsAdded += dimResult.DimensionsAdded;
                            result.ViewsCreated++;

                            ErrorHandler.DebugLog($"{proc}: Created TopView '{topView.GetName2()}', dims={dimResult.DimensionsAdded}");
                        }
                        else
                        {
                            ErrorHandler.DebugLog($"{proc}: CreateUnfoldedViewAt3 for TopView returned null");
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorHandler.DebugLog($"{proc}: TopView error: {ex.Message}");
                    }
                }

                // RightView: created if horizontal bends exist (shows right-side formed shape)
                if (hasHorizontalBends)
                {
                    try
                    {
                        drawDoc.ActivateView(flatViewName);
                        drawModel.Extension.SelectByID2(flatViewName, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

                        // VBA: CreateUnfoldedViewAt3(0.3, flatPos.Y, 0, false)
                        var rightView = drawDoc.CreateUnfoldedViewAt3(ProjectedOffset, flatPos[1], 0, false) as IView;
                        if (rightView != null)
                        {
                            drawModel.ClearSelection2(true);
                            rightView.ReferencedConfiguration = "Default";
                            rightView.SetDisplayMode3(false, 4 /* swHIDDEN_GREYED */, false, true);
                            drawModel.ForceRebuild3(false);

                            // Clamp within sheet bounds
                            ClampViewToSheet(drawDoc, rightView);

                            // Dimension with overall W/H only
                            var dimResult = dimensioner.DimensionFormedView(drawDoc, rightView);
                            result.DimensionsAdded += dimResult.DimensionsAdded;
                            result.ViewsCreated++;

                            ErrorHandler.DebugLog($"{proc}: Created RightView '{rightView.GetName2()}', dims={dimResult.DimensionsAdded}");
                        }
                        else
                        {
                            ErrorHandler.DebugLog($"{proc}: CreateUnfoldedViewAt3 for RightView returned null");
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorHandler.DebugLog($"{proc}: RightView error: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"{proc}: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Clamps a view within the sheet bounds (MaxRight, MaxTop, LeftMargin, BottomMargin).
        /// </summary>
        private void ClampViewToSheet(IDrawingDoc drawDoc, IView view)
        {
            try
            {
                var drawModel = drawDoc as IModelDoc2;
                var outline = view.GetOutline() as double[];
                var pos = view.Position as double[];
                if (outline == null || pos == null) return;

                bool moved = false;
                if (outline[2] > MaxRight)  { pos[0] -= (outline[2] - MaxRight); moved = true; }
                if (outline[3] > MaxTop)    { pos[1] -= (outline[3] - MaxTop); moved = true; }
                if (outline[0] < LeftMargin) { pos[0] += (LeftMargin - outline[0]); moved = true; }
                if (outline[1] < BottomMargin) { pos[1] += (BottomMargin - outline[1]); moved = true; }

                if (moved)
                {
                    view.Position = pos;
                    if (drawModel != null) drawModel.EditRebuild3();
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"ClampViewToSheet: {ex.Message}");
            }
        }

        /// <summary>
        /// Scales views to fit the sheet. First scales DOWN if views overflow, then
        /// scales UP if views are too small (VBA behavior). Uses "use scale" on each
        /// view rather than changing the sheet scale, since view scale is more reliable
        /// for mixed view sizes.
        /// </summary>
        private void AutoScaleToFit(IDrawingDoc drawDoc, IModelDoc2 drawModel)
        {
            try
            {
                var sheet = drawDoc.GetCurrentSheet() as ISheet;
                if (sheet == null) return;

                var viewsRaw = sheet.GetViews() as object[];
                if (viewsRaw == null || viewsRaw.Length == 0) return;

                // Get sheet dimensions in meters
                double sheetW = 0, sheetH = 0;
                sheet.GetSize(ref sheetW, ref sheetH);
                if (sheetW <= 0 || sheetH <= 0) return;

                // Find the overall bounding box of all views
                double allLeft = double.MaxValue, allBottom = double.MaxValue;
                double allRight = double.MinValue, allTop = double.MinValue;
                foreach (var viewObj in viewsRaw)
                {
                    var view = viewObj as IView;
                    if (view == null) continue;
                    var outline = view.GetOutline() as double[];
                    if (outline == null || outline.Length < 4) continue;
                    if (outline[0] < allLeft) allLeft = outline[0];
                    if (outline[1] < allBottom) allBottom = outline[1];
                    if (outline[2] > allRight) allRight = outline[2];
                    if (outline[3] > allTop) allTop = outline[3];
                }

                double totalW = allRight - allLeft;
                double totalH = allTop - allBottom;
                double usableW = MaxRight - LeftMargin;
                double usableH = MaxTop - BottomMargin;

                // If views already fit, nothing to do
                if (totalW <= usableW && totalH <= usableH)
                    return;

                // Views overflow: compute required scale reduction
                double scaleX = usableW / totalW;
                double scaleY = usableH / totalH;
                double scaleFactor = Math.Min(scaleX, scaleY) * 0.95; // 5% margin

                if (scaleFactor >= 1.0) return; // No reduction needed

                ErrorHandler.DebugLog($"[DWG] AutoScaleDown: views {totalW * MetersToIn:F2}\"x{totalH * MetersToIn:F2}\" " +
                    $"vs usable {usableW * MetersToIn:F2}\"x{usableH * MetersToIn:F2}\", scale={scaleFactor:F3}");

                // Apply scale reduction to each view
                foreach (var viewObj in viewsRaw)
                {
                    var view = viewObj as IView;
                    if (view == null) continue;

                    var scaleRatio = view.ScaleRatio as double[];
                    if (scaleRatio == null || scaleRatio.Length < 2) continue;

                    view.ScaleRatio = new double[] { scaleRatio[0] * scaleFactor, scaleRatio[1] };
                }

                drawModel.ForceRebuild3(false);

                // Reposition all views within sheet bounds
                foreach (var viewObj in viewsRaw)
                {
                    var view = viewObj as IView;
                    if (view == null) continue;
                    ClampViewToSheet(drawDoc, view);
                }

                // Reposition primary view (first view) to margins
                var firstView = viewsRaw[0] as IView;
                if (firstView != null)
                {
                    var outline = firstView.GetOutline() as double[];
                    var pos = firstView.Position as double[];
                    if (outline != null && pos != null)
                    {
                        pos[0] += LeftMargin - outline[0];
                        pos[1] += BottomMargin - outline[1];
                        firstView.Position = pos;
                        drawModel.EditRebuild3();
                    }
                }

                // Reposition secondary views relative to primary
                if (viewsRaw.Length > 1 && firstView != null)
                {
                    var primaryOutline = firstView.GetOutline() as double[];
                    for (int i = 1; i < viewsRaw.Length; i++)
                    {
                        var secView = viewsRaw[i] as IView;
                        if (secView == null || primaryOutline == null) continue;

                        var secOutline = secView.GetOutline() as double[];
                        var secPos = secView.Position as double[];
                        if (secOutline == null || secPos == null) continue;

                        // If secondary view overlaps or is off-sheet, reposition
                        if (secOutline[0] < primaryOutline[2] + ViewGap)
                        {
                            secPos[0] += (primaryOutline[2] + ViewGap) - secOutline[0];
                            secView.Position = secPos;
                        }
                        ClampViewToSheet(drawDoc, secView);
                    }
                    drawModel.EditRebuild3();
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"AutoScaleToFit: {ex.Message}");
            }
        }

        /// <summary>
        /// Scales sheet numerator up by 5% per iteration to fill available space (VBA behavior).
        /// Only used when views are already within bounds but have extra space.
        /// </summary>
        private void ScaleSheetUp(ISheet sheet, IDrawingDoc drawDoc, IModelDoc2 drawModel)
        {
            try
            {
                for (int iter = 0; iter < ScaleMaxIter; iter++)
                {
                    var viewsRaw = sheet.GetViews() as object[];
                    if (viewsRaw == null) break;

                    bool allFit = true;
                    foreach (var viewObj in viewsRaw)
                    {
                        var view = viewObj as IView;
                        if (view == null) continue;
                        var outline = view.GetOutline() as double[];
                        if (outline == null) continue;
                        if (outline[2] > MaxRight || outline[3] > MaxTop)
                        {
                            allFit = false;
                            break;
                        }
                    }
                    if (!allFit) break;

                    var props = sheet.GetProperties2() as double[];
                    if (props == null || props.Length < 8) break;

                    props[2] *= ScaleFactor;
                    sheet.SetProperties2(
                        (int)props[0], (int)props[1], props[2], props[3],
                        props[4] != 0, props[5], props[6], props[7] != 0);

                    drawModel.EditRebuild3();
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"ScaleSheetUp: {ex.Message}");
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

        /// <summary>
        /// Hides standard reference planes (Front, Top, Right) so they don't clutter drawing views.
        /// Ported from VBA SP.bas lines 1883-1899.
        /// </summary>
        private int HideReferencePlanes(IModelDoc2 model)
        {
            int count = 0;
            try
            {
                var feat = model.FirstFeature() as IFeature;
                while (feat != null)
                {
                    string typeName = feat.GetTypeName2() ?? "";
                    if (typeName.Equals("RefPlane", StringComparison.OrdinalIgnoreCase))
                    {
                        feat.Select2(true, 0);
                        model.BlankRefGeom();
                        count++;
                    }
                    feat = feat.GetNextFeature() as IFeature;
                }
                model.ClearSelection2(true);
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"HideReferencePlanes: {ex.Message}");
            }
            return count;
        }

        /// <summary>
        /// Makes ProfileFeature sketches (laser etch marks) visible in the drawing view.
        /// These are hidden by default but need to appear on shop prints.
        /// Ported from VBA SP.bas lines 1911-1951.
        /// </summary>
        private int MakeEtchMarksVisible(IDrawingDoc drawDoc, IView view, IModelDoc2 partModel)
        {
            int count = 0;
            try
            {
                if (drawDoc == null || view == null || partModel == null) return 0;

                var drawModel = drawDoc as IModelDoc2;
                string viewUniqueName = view.GetName2();
                string modelTitle = partModel.GetTitle();
                if (string.IsNullOrEmpty(modelTitle)) return 0;

                // Strip file extension from model title if present
                if (modelTitle.EndsWith(".SLDPRT", StringComparison.OrdinalIgnoreCase) ||
                    modelTitle.EndsWith(".sldprt", StringComparison.OrdinalIgnoreCase))
                    modelTitle = modelTitle.Substring(0, modelTitle.Length - 7);

                // Switch part to the view's referenced configuration
                string viewConfig = view.ReferencedConfiguration;
                if (!string.IsNullOrEmpty(viewConfig))
                    partModel.ShowConfiguration2(viewConfig);

                // Iterate features to find hidden ProfileFeature sketches
                var feat = partModel.FirstFeature() as IFeature;
                while (feat != null)
                {
                    string typeName = feat.GetTypeName2() ?? "";
                    if (typeName.Equals("ProfileFeature", StringComparison.OrdinalIgnoreCase))
                    {
                        string sketchName = feat.Name ?? "";

                        // Skip "Bounding Box" type sketches (VBA: skip names starting with "Bound")
                        if (!sketchName.StartsWith("Bound", StringComparison.OrdinalIgnoreCase))
                        {
                            // Check if sketch is hidden (Visible == 2 means hidden)
                            int visible = feat.Visible;
                            if (visible == 2)
                            {
                                // Build selection name: "sketchName@modelTitle@viewUniqueName"
                                string selName = $"{sketchName}@{modelTitle}@{viewUniqueName}";

                                // Activate the drawing view first
                                drawDoc.ActivateView(viewUniqueName);

                                // Select the sketch in the drawing view context
                                bool selected = drawModel.Extension.SelectByID2(
                                    selName, "SKETCH", 0, 0, 0, false, 0, null, 0);

                                if (selected)
                                {
                                    // UnblankSketch makes the sketch visible in the drawing
                                    drawModel.UnblankSketch();
                                    count++;
                                    ErrorHandler.DebugLog($"[DWG] Etch mark visible: '{sketchName}'");
                                }
                            }
                        }
                    }
                    feat = feat.GetNextFeature() as IFeature;
                }

                drawModel.ClearSelection2(true);

                // Switch back to Default config
                partModel.ShowConfiguration2("Default");
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"MakeEtchMarksVisible: {ex.Message}");
            }
            return count;
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
