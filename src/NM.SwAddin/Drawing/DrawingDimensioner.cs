using System;
using System.Collections.Generic;
using NM.Core;
using NM.Core.Drawing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Drawing
{
    /// <summary>
    /// Auto-dimensions drawing views: flat pattern bend-to-bend dimensions,
    /// overall dimensions, tube profile dimensions, and formed view dimensions.
    /// Ported from VBA DimensionDrawing.bas: DimensionFlat, DimensionTube,
    /// DimensionOther, RectProfile, and AlignDims.bas: Align.
    /// </summary>
    public sealed class DrawingDimensioner
    {
        private readonly ISldWorks _swApp;
        private readonly ViewGeometryAnalyzer _analyzer;

        /// <summary>
        /// Conversion factor: inches per meter (used in VBA as 39.3700787401575).
        /// </summary>
        private const double InchesPerMeter = 39.3700787401575;

        public DrawingDimensioner(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
            _analyzer = new ViewGeometryAnalyzer(swApp);
        }

        /// <summary>
        /// Result of a dimensioning operation.
        /// </summary>
        public sealed class DimensionResult
        {
            public bool Success { get; set; }
            public int DimensionsAdded { get; set; }
            public int ViewsCreated { get; set; }
            public string Message { get; set; }

            /// <summary>Whether vertical bend lines were found (need TopView projected).</summary>
            public bool HasVerticalBends { get; set; }

            /// <summary>Whether horizontal bend lines were found (need RightView projected).</summary>
            public bool HasHorizontalBends { get; set; }
        }

        /// <summary>
        /// Dimensions a flat pattern view: overall width/height, bend-to-bend dimensions
        /// with " BL" suffix.
        /// Ported from VBA DimensionFlat() (lines 44-463).
        /// </summary>
        public DimensionResult DimensionFlatPattern(IDrawingDoc drawDoc, IView view)
        {
            const string proc = nameof(DimensionFlatPattern);
            ErrorHandler.PushCallStack(proc);

            var result = new DimensionResult();
            try
            {
                var swModel = drawDoc as IModelDoc2;
                if (swModel == null)
                {
                    result.Message = "Cannot cast drawDoc to IModelDoc2";
                    return result;
                }

                // Get view scale
                var origVXf = view.GetXform() as double[];
                if (origVXf == null || origVXf.Length < 3)
                {
                    result.Message = "Cannot get view transform";
                    return result;
                }
                double viewScale = origVXf[2];
                if (viewScale == 0) viewScale = 1;

                // Step 1: Find bend lines
                List<BendElement> horzBends, vertBends;
                _analyzer.FindBendLines(drawDoc, view, out horzBends, out vertBends);
                ErrorHandler.DebugLog($"{proc}: Found {horzBends.Count} horizontal bends, {vertBends.Count} vertical bends");

                // Step 2: Find boundary edges with retry (VBA retries up to 25 times)
                EdgeElement leftPos = null, rightPos = null, topPos = null, bottomPos = null;
                for (int attempt = 0; attempt < 25 && leftPos == null; attempt++)
                    leftPos = _analyzer.FindBoundaryEdge(view, drawDoc, BoundaryDirection.Left);
                for (int attempt = 0; attempt < 25 && rightPos == null; attempt++)
                    rightPos = _analyzer.FindBoundaryEdge(view, drawDoc, BoundaryDirection.Right);
                for (int attempt = 0; attempt < 25 && topPos == null; attempt++)
                    topPos = _analyzer.FindBoundaryEdge(view, drawDoc, BoundaryDirection.Top);
                for (int attempt = 0; attempt < 25 && bottomPos == null; attempt++)
                    bottomPos = _analyzer.FindBoundaryEdge(view, drawDoc, BoundaryDirection.Bottom);

                // If boundaries missing, fall back to tube dimensioning
                if (leftPos == null || rightPos == null || topPos == null || bottomPos == null)
                {
                    ErrorHandler.DebugLog($"{proc}: Boundary edges missing, falling back to DimensionTube");
                    return DimensionTube(drawDoc, view);
                }

                ErrorHandler.DebugLog($"{proc}: Boundaries L={leftPos.X:F3} R={rightPos.X:F3} T={topPos.Y:F3} B={bottomPos.Y:F3}");

                drawDoc.ActivateView(view.GetName2());

                // Step 3: Overall horizontal dimension (left → right)
                int dimsBefore = result.DimensionsAdded;
                try
                {
                    swModel.ClearSelection2(true);
                    SelectObject(leftPos.Obj, true);
                    SelectObject(rightPos.Obj, true);

                    // Dim placement: centered between left/right, below bottom edge
                    double horzDimX = ((rightPos.X - leftPos.X) / 2.0 + leftPos.X) * viewScale / InchesPerMeter;
                    double horzDimY = (bottomPos.Y * viewScale - 0.25) / InchesPerMeter;
                    var dim = swModel.AddHorizontalDimension2(horzDimX, horzDimY, 0);
                    if (dim != null) result.DimensionsAdded++;

                    swModel.ClearSelection2(true);
                    swModel.EditRebuild3();
                }
                catch (Exception ex) { ErrorHandler.DebugLog($"{proc}: Horz overall dim: {ex.Message}"); }

                // Step 4: Overall vertical dimension (top → bottom)
                try
                {
                    swModel.ClearSelection2(true);
                    drawDoc.ActivateView(view.GetName2());
                    SelectObject(topPos.Obj, true);
                    SelectObject(bottomPos.Obj, true);

                    double vertDimX = (leftPos.X * viewScale - 0.25) / InchesPerMeter;
                    double vertDimY = ((topPos.Y - bottomPos.Y) / 2.0 + bottomPos.Y) * viewScale / InchesPerMeter;
                    var dim = swModel.AddVerticalDimension2(vertDimX, vertDimY, 0);
                    if (dim != null) result.DimensionsAdded++;

                    swModel.ClearSelection2(true);
                    swModel.EditRebuild3();
                }
                catch (Exception ex) { ErrorHandler.DebugLog($"{proc}: Vert overall dim: {ex.Message}"); }

                // Step 5: Vertical bends → horizontal bend-to-bend dims with " BL" suffix
                result.HasVerticalBends = vertBends.Count > 0;
                if (vertBends.Count > 0)
                {
                    result.DimensionsAdded += DimensionBendLines(
                        swModel, drawDoc, view, vertBends, leftPos, rightPos, bottomPos,
                        viewScale, isVerticalBends: true);
                }

                // Step 6: Horizontal bends → vertical bend-to-bend dims with " BL" suffix
                result.HasHorizontalBends = horzBends.Count > 0;
                if (horzBends.Count > 0)
                {
                    result.DimensionsAdded += DimensionBendLines(
                        swModel, drawDoc, view, horzBends, bottomPos, topPos, leftPos,
                        viewScale, isVerticalBends: false);
                }

                result.Success = true;
                result.Message = $"Added {result.DimensionsAdded} dimensions";
                ErrorHandler.DebugLog($"{proc}: {result.Message}");
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
        /// Creates bend-to-bend dimensions between adjacent bends (and from boundary to first/last bend).
        /// Adds " BL" suffix to each dimension.
        /// </summary>
        /// <param name="bends">Sorted bend elements.</param>
        /// <param name="startEdge">Starting boundary edge (Left for vertical bends, Bottom for horizontal).</param>
        /// <param name="endEdge">Ending boundary edge (Right for vertical bends, Top for horizontal).</param>
        /// <param name="offsetEdge">Edge used for dim offset (Bottom for vertical bends, Left for horizontal).</param>
        /// <param name="isVerticalBends">True for vertical bends (horizontal dims), false for horizontal bends (vertical dims).</param>
        /// <returns>Number of dimensions added.</returns>
        private int DimensionBendLines(IModelDoc2 swModel, IDrawingDoc drawDoc, IView view,
            List<BendElement> bends, EdgeElement startEdge, EdgeElement endEdge, EdgeElement offsetEdge,
            double viewScale, bool isVerticalBends)
        {
            int dimsAdded = 0;
            string viewName = view.GetName2();

            // Get root drawing component name for selection string
            string compName = "";
            try
            {
                var rootComp = view.RootDrawingComponent;
                if (rootComp != null)
                    compName = rootComp.Name ?? "";
            }
            catch { }

            // Build unique bend positions (skip co-located bends that have the same Position)
            var uniqueBends = new List<BendElement>();
            double lastPosition = double.NaN;
            foreach (var bend in bends)
            {
                if (double.IsNaN(lastPosition) || Math.Abs(bend.Position - lastPosition) > 0.001)
                {
                    uniqueBends.Add(bend);
                    lastPosition = bend.Position;
                }
            }

            if (uniqueBends.Count == 0) return 0;

            // Dim between start boundary → first bend → ... → last bend → end boundary
            for (int i = 0; i <= uniqueBends.Count; i++)
            {
                try
                {
                    swModel.ClearSelection2(true);
                    drawDoc.ActivateView(viewName);

                    double centerLoc;

                    if (i == 0)
                    {
                        // Start boundary → first bend
                        double startPos = isVerticalBends ? startEdge.X : startEdge.Y;
                        double bendPos = uniqueBends[0].Position;
                        centerLoc = ((bendPos - startPos) / 2.0 + startPos) * viewScale / InchesPerMeter;

                        // Select start boundary edge
                        SelectObject(startEdge.Obj, true);
                        // Select first bend line by name
                        SelectBendLine(swModel, uniqueBends[0], compName, viewName);
                    }
                    else if (i == uniqueBends.Count)
                    {
                        // Last bend → end boundary
                        double bendPos = uniqueBends[i - 1].Position;
                        double endPos = isVerticalBends ? endEdge.X : endEdge.Y;
                        centerLoc = ((endPos - bendPos) / 2.0 + bendPos) * viewScale / InchesPerMeter;

                        // Select last bend line
                        SelectBendLine(swModel, uniqueBends[i - 1], compName, viewName);
                        // Select end boundary edge
                        SelectObject(endEdge.Obj, true);
                    }
                    else
                    {
                        // Bend i-1 → Bend i
                        double prevPos = uniqueBends[i - 1].Position;
                        double currPos = uniqueBends[i].Position;
                        centerLoc = ((currPos - prevPos) / 2.0 + prevPos) * viewScale / InchesPerMeter;

                        SelectBendLine(swModel, uniqueBends[i - 1], compName, viewName);
                        SelectBendLine(swModel, uniqueBends[i], compName, viewName);
                    }

                    // Add the dimension
                    object dim = null;
                    if (isVerticalBends)
                    {
                        // Horizontal dim: Y offset below bottom
                        double dimY = (offsetEdge.Y * viewScale - 0.125) / InchesPerMeter;
                        dim = swModel.AddHorizontalDimension2(centerLoc, dimY, 0);
                    }
                    else
                    {
                        // Vertical dim: X offset left of left
                        double dimX = (offsetEdge.X * viewScale - 0.125) / InchesPerMeter;
                        dim = swModel.AddVerticalDimension2(dimX, centerLoc, 0);
                    }

                    if (dim != null)
                    {
                        // Add " BL" suffix
                        // EditDimensionProperties2(precision, primary tol, secondary tol,
                        //   prefix, suffix, showTol, tolType, secondDisplay, showParens,
                        //   toleranceFit_Upper, toleranceFit_Lower, primaryText, dualText,
                        //   showDual, dualPrefix, dualSuffix, bendsReverse)
                        swModel.EditDimensionProperties2(0, 0, 0, "", "", true, 9, 2, true, 12, 12, "", " BL", true, "", "", false);
                        dimsAdded++;
                    }

                    swModel.ClearSelection2(true);
                    swModel.EditRebuild3();
                }
                catch (Exception ex)
                {
                    ErrorHandler.DebugLog($"DimensionBendLines[{i}]: {ex.Message}");
                }
            }

            return dimsAdded;
        }

        /// <summary>
        /// Selects a bend line by its fully-qualified name in the drawing view.
        /// Name format: "{segName}@{sketchName}@{componentName}@{viewName}"
        /// </summary>
        private bool SelectBendLine(IModelDoc2 swModel, BendElement bend, string compName, string viewName)
        {
            try
            {
                var seg = bend.SketchSegment as ISketchSegment;
                if (seg == null) return false;

                string segName = seg.GetName() ?? "";
                var sketch = seg.GetSketch() as ISketch;
                string sketchName = "";
                if (sketch != null)
                {
                    var feat = sketch as IFeature;
                    if (feat != null)
                        sketchName = feat.Name ?? "";
                }

                string fullName = $"{segName}@{sketchName}@{compName}@{viewName}";
                return swModel.Extension.SelectByID2(fullName, "EXTSKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"SelectBendLine: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Selects a SolidWorks object (edge, vertex, etc.) using IEntity.Select4.
        /// All SW geometry objects implement IEntity which provides Select4.
        /// </summary>
        private bool SelectObject(object obj, bool append)
        {
            if (obj == null) return false;
            try
            {
                // Cast to IEntity — edges, vertices, sketch segments all implement it
                var entity = obj as IEntity;
                if (entity != null)
                    return entity.Select4(append, null);

                // Fallback: try via reflection for COM objects that may not cast cleanly
                var method = obj.GetType().GetMethod("Select4");
                if (method != null)
                    return (bool)method.Invoke(obj, new object[] { append, null });
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"SelectObject: {ex.Message}");
            }
            return false;
        }

        /// <summary>
        /// Dimensions a tube view: detects circular or rectangular profile,
        /// adds diameter or width/height dimensions.
        /// Ported from VBA DimensionTube() (lines 572-643).
        /// </summary>
        public DimensionResult DimensionTube(IDrawingDoc drawDoc, IView view)
        {
            const string proc = nameof(DimensionTube);
            ErrorHandler.PushCallStack(proc);

            var result = new DimensionResult();
            try
            {
                var swModel = drawDoc as IModelDoc2;
                if (swModel == null)
                {
                    result.Message = "Cannot cast drawDoc to IModelDoc2";
                    return result;
                }

                var origVXf = view.GetXform() as double[];
                if (origVXf == null || origVXf.Length < 3)
                {
                    result.Message = "Cannot get view transform";
                    return result;
                }
                double viewScale = origVXf[2];
                if (viewScale == 0) viewScale = 1;

                // Get view outline
                var outline = view.GetOutline() as double[];
                if (outline == null || outline.Length < 4)
                {
                    result.Message = "Cannot get view outline";
                    return result;
                }

                drawDoc.ActivateView(view.GetName2());

                // Calculate edge positions from view outline (VBA lines 590-597)
                double leftX = outline[0] * InchesPerMeter / viewScale;
                double rightX = outline[2] * InchesPerMeter / viewScale;
                double bottomY = outline[1] * InchesPerMeter / viewScale;
                double topY = outline[3] * InchesPerMeter / viewScale;
                double midY = ((outline[3] - outline[1]) / 2.0 + outline[1]) * InchesPerMeter / viewScale;
                double midX = ((outline[2] - outline[0]) / 2.0 + outline[0]) * InchesPerMeter / viewScale;

                // Try GetVisibleEntities2 to find circular edges (reliable for round tubes)
                try
                {
                    Component2 comp = null;
                    try
                    {
                        var rootDrawComp = view.RootDrawingComponent;
                        if (rootDrawComp != null)
                            comp = rootDrawComp.Component as Component2;
                    }
                    catch { }

                    // swViewEntityType_Edge = 1
                    var edgesRaw = view.GetVisibleEntities2(comp, 1) as object[];
                    ErrorHandler.DebugLog($"{proc}: GetVisibleEntities2 returned {edgesRaw?.Length ?? 0} edges");

                    if (edgesRaw != null && edgesRaw.Length > 0)
                    {
                        IEdge bestCircle = null;
                        double bestRadius = 0;

                        foreach (var edgeObj in edgesRaw)
                        {
                            var edge = edgeObj as IEdge;
                            if (edge == null) continue;

                            var curve = edge.GetCurve() as ICurve;
                            if (curve == null) continue;

                            if (curve.IsCircle())
                            {
                                var circParams = curve.CircleParams as double[];
                                if (circParams != null && circParams.Length >= 7)
                                {
                                    double radius = circParams[6];
                                    if (radius > bestRadius)
                                    {
                                        bestRadius = radius;
                                        bestCircle = edge;
                                    }
                                }
                            }
                        }

                        if (bestCircle != null)
                        {
                            ErrorHandler.DebugLog($"{proc}: Found circular edge, OD={bestRadius * 2 * InchesPerMeter:F3}\"");
                            swModel.ClearSelection2(true);

                            var entity = bestCircle as IEntity;
                            if (entity != null && entity.Select4(true, null))
                            {
                                double dimX = (outline[0] + outline[2]) / 2.0;
                                double dimY = outline[1];
                                var dim = swModel.AddDiameterDimension2(dimX, dimY, 0);
                                if (dim != null)
                                {
                                    result.DimensionsAdded++;
                                    ErrorHandler.DebugLog($"{proc}: Added diameter dimension");
                                }
                            }
                            swModel.ClearSelection2(true);

                            if (result.DimensionsAdded > 0)
                            {
                                result.Success = true;
                                result.Message = $"Added {result.DimensionsAdded} tube dimensions";
                                return result;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandler.DebugLog($"{proc}: GetVisibleEntities2 approach: {ex.Message}");
                }

                // Fallback: try SelectByRay at left edge (works for some profiles)
                swModel.ClearSelection2(true);
                double rayX = leftX * 25.4 / 1000.0 * viewScale;
                double rayY = midY * 25.4 / 1000.0 * viewScale;
                bool selected = swModel.Extension.SelectByRay(rayX, rayY, -500, 0, 0, -1, 0.005, 1, true, 0, 0);

                var selMgr = swModel.SelectionManager as ISelectionMgr;
                if (selected && selMgr != null && selMgr.GetSelectedObjectCount2(-1) > 0)
                {
                    // Use SketchUseEdge3 to convert to sketch, then check if arc
                    swModel.SketchManager.SketchUseEdge3(false, false);

                    bool isArc = false;
                    try
                    {
                        var selObj = selMgr.GetSelectedObject6(1, -1);
                        if (selObj != null)
                        {
                            var skSeg = selObj as ISketchSegment;
                            if (skSeg != null && skSeg.GetType() == (int)swSketchSegments_e.swSketchARC)
                                isArc = true;
                        }
                    }
                    catch { }

                    if (isArc)
                    {
                        swModel.EditUndo2(1);
                        swModel.ClearSelection2(true);
                        swModel.Extension.SelectByRay(rayX, rayY, -500, 0, 0, -1, 0.005, 1, true, 0, 0);
                        var dim = swModel.AddDiameterDimension2(outline[0], outline[1], 0);
                        if (dim != null) result.DimensionsAdded++;
                        swModel.ClearSelection2(true);
                    }
                    else
                    {
                        swModel.EditUndo2(1);
                        swModel.ClearSelection2(true);
                        selected = false;
                    }
                }

                // If not circular, try rectangular profile
                if (result.DimensionsAdded == 0)
                {
                    var leftElem = new EdgeElement { X = leftX, Y = midY };
                    var rightElem = new EdgeElement { X = rightX, Y = midY };
                    var topElem = new EdgeElement { X = midX, Y = topY };
                    var bottomElem = new EdgeElement { X = midX, Y = bottomY };

                    if (DimensionRectProfile(drawDoc, view, leftElem, rightElem, topElem, bottomElem))
                    {
                        result.DimensionsAdded += 2;
                    }
                    else
                    {
                        // Fallback: try bottom edge for a horizontal dim
                        try
                        {
                            double bRayX = midX * 25.4 / 1000.0 * viewScale;
                            double bRayY = bottomY * 25.4 / 1000.0 * viewScale;
                            swModel.ClearSelection2(true);
                            bool b = swModel.Extension.SelectByRay(bRayX, bRayY, -500, 0, 0, -1, 0.005, 1, true, 0, 0);
                            if (b)
                            {
                                var dim = swModel.AddHorizontalDimension2(midX, bottomY, 0);
                                if (dim != null) result.DimensionsAdded++;
                            }
                        }
                        catch { }
                    }
                }

                result.Success = result.DimensionsAdded > 0;
                result.Message = result.DimensionsAdded > 0
                    ? $"Added {result.DimensionsAdded} tube dimensions"
                    : "No tube dimensions added";
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
        /// Dimensions a formed (3D) projected view with overall width and height.
        /// Ported from VBA DimensionOther() (lines 466-507).
        /// </summary>
        public DimensionResult DimensionFormedView(IDrawingDoc drawDoc, IView view)
        {
            const string proc = nameof(DimensionFormedView);
            ErrorHandler.PushCallStack(proc);

            var result = new DimensionResult();
            try
            {
                var swModel = drawDoc as IModelDoc2;
                if (swModel == null)
                {
                    result.Message = "Cannot cast drawDoc to IModelDoc2";
                    return result;
                }

                var origVXf = view.GetXform() as double[];
                if (origVXf == null || origVXf.Length < 3)
                {
                    result.Message = "Cannot get view transform";
                    return result;
                }
                double viewScale = origVXf[2];
                if (viewScale == 0) viewScale = 1;

                // Find boundary edges
                var leftPos = _analyzer.FindBoundaryEdge(view, drawDoc, BoundaryDirection.Left);
                var rightPos = _analyzer.FindBoundaryEdge(view, drawDoc, BoundaryDirection.Right);
                var topPos = _analyzer.FindBoundaryEdge(view, drawDoc, BoundaryDirection.Top);
                var bottomPos = _analyzer.FindBoundaryEdge(view, drawDoc, BoundaryDirection.Bottom);

                if (leftPos == null || rightPos == null || topPos == null || bottomPos == null)
                {
                    // Fallback: use view outline + SelectByRay (for tube side views where
                    // boundary edges are circular/silhouette and can't be found as lines)
                    ErrorHandler.DebugLog($"{proc}: Missing boundaries L={leftPos != null} R={rightPos != null} T={topPos != null} B={bottomPos != null}, using outline fallback");
                    var outline = view.GetOutline() as double[];
                    if (outline != null && outline.Length >= 4)
                    {
                        double oLeftX = outline[0] * InchesPerMeter / viewScale;
                        double oRightX = outline[2] * InchesPerMeter / viewScale;
                        double oBottomY = outline[1] * InchesPerMeter / viewScale;
                        double oTopY = outline[3] * InchesPerMeter / viewScale;
                        double oMidY = ((outline[3] - outline[1]) / 2.0 + outline[1]) * InchesPerMeter / viewScale;
                        double oMidX = ((outline[2] - outline[0]) / 2.0 + outline[0]) * InchesPerMeter / viewScale;

                        var lElem = new EdgeElement { X = oLeftX, Y = oMidY };
                        var rElem = new EdgeElement { X = oRightX, Y = oMidY };
                        var tElem = new EdgeElement { X = oMidX, Y = oTopY };
                        var bElem = new EdgeElement { X = oMidX, Y = oBottomY };

                        if (DimensionRectProfile(drawDoc, view, lElem, rElem, tElem, bElem))
                        {
                            result.DimensionsAdded += 2;
                            result.Success = true;
                            result.Message = "Added 2 formed view dimensions (outline fallback)";
                        }
                        else
                        {
                            result.Message = "Cannot find boundary edges or select edges for formed view";
                        }
                    }
                    else
                    {
                        result.Message = "Cannot find boundary edges or outline for formed view";
                    }
                    return result;
                }

                drawDoc.ActivateView(view.GetName2());

                // Horizontal dimension: left → right
                try
                {
                    swModel.ClearSelection2(true);
                    SelectObject(leftPos.Obj, true);
                    SelectObject(rightPos.Obj, true);
                    // VBA uses 0.5 offset for DimensionOther (vs 0.25 for DimensionFlat)
                    double horzX = ((rightPos.X - leftPos.X) / 2.0 + leftPos.X) * viewScale / InchesPerMeter;
                    double horzY = (bottomPos.Y * viewScale - 0.5) / InchesPerMeter;
                    var dim = swModel.AddHorizontalDimension2(horzX, horzY, 0);
                    if (dim != null) result.DimensionsAdded++;
                    swModel.ClearSelection2(true);
                }
                catch (Exception ex) { ErrorHandler.DebugLog($"{proc}: Horz dim: {ex.Message}"); }

                // Vertical dimension: top → bottom
                try
                {
                    swModel.ClearSelection2(true);
                    SelectObject(topPos.Obj, true);
                    SelectObject(bottomPos.Obj, true);
                    double vertX = (leftPos.X * viewScale - 0.5) / InchesPerMeter;
                    double vertY = ((topPos.Y - bottomPos.Y) / 2.0 + bottomPos.Y) * viewScale / InchesPerMeter;
                    var dim = swModel.AddVerticalDimension2(vertX, vertY, 0);
                    if (dim != null) result.DimensionsAdded++;
                    swModel.ClearSelection2(true);
                }
                catch (Exception ex) { ErrorHandler.DebugLog($"{proc}: Vert dim: {ex.Message}"); }

                result.Success = result.DimensionsAdded > 0;
                result.Message = $"Added {result.DimensionsAdded} formed view dimensions";
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
        /// Attempts to dimension a rectangular tube profile by selecting opposing edges via SelectByRay.
        /// Ported from VBA RectProfile() (lines 645-676).
        /// </summary>
        public bool DimensionRectProfile(IDrawingDoc drawDoc, IView view,
            EdgeElement left, EdgeElement right, EdgeElement top, EdgeElement bottom)
        {
            try
            {
                var swModel = drawDoc as IModelDoc2;
                if (swModel == null) return false;

                var origVXf = view.GetXform() as double[];
                if (origVXf == null || origVXf.Length < 3) return false;
                double viewScale = origVXf[2];
                if (viewScale == 0) viewScale = 1;

                // Select left edge via ray
                double lRayX = left.X * 25.4 / 1000.0 * viewScale;
                double lRayY = left.Y * 25.4 / 1000.0 * viewScale;
                bool b = swModel.Extension.SelectByRay(lRayX, lRayY, -500, 0, 0, -1, 0.005, 1, true, 0, 0);
                if (!b) return false;

                // Select right edge via ray (append)
                double rRayX = right.X * 25.4 / 1000.0 * viewScale;
                double rRayY = right.Y * 25.4 / 1000.0 * viewScale;
                b = swModel.Extension.SelectByRay(rRayX, rRayY, -500, 0, 0, -1, 0.005, 1, true, 0, 0);
                if (!b) return false;

                // Add horizontal dimension
                double horzX = ((right.X - left.X) / 2.0 + left.X) * viewScale / InchesPerMeter;
                double horzY = (bottom.Y * viewScale - 0.25) / InchesPerMeter;
                var hDim = swModel.AddHorizontalDimension2(horzX, horzY, 0);
                if (hDim == null) return false;

                // Select top edge via ray
                double tRayX = top.X * 25.4 / 1000.0 * viewScale;
                double tRayY = top.Y * 25.4 / 1000.0 * viewScale;
                b = swModel.Extension.SelectByRay(tRayX, tRayY, -500, 0, 0, -1, 0.005, 1, true, 0, 0);
                if (!b) return false;

                // Select bottom edge via ray (append)
                double bRayX = bottom.X * 25.4 / 1000.0 * viewScale;
                double bRayY = bottom.Y * 25.4 / 1000.0 * viewScale;
                b = swModel.Extension.SelectByRay(bRayX, bRayY, -500, 0, 0, -1, 0.005, 1, true, 0, 0);
                if (!b) return false;

                // Add vertical dimension
                double vertX = (left.X * viewScale - 0.25) / InchesPerMeter;
                double vertY = ((top.Y - bottom.Y) / 2.0 + bottom.Y) * viewScale / InchesPerMeter;
                var vDim = swModel.AddVerticalDimension2(vertX, vertY, 0);
                if (vDim == null) return false;

                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"DimensionRectProfile: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Auto-arranges all dimensions in the drawing using SolidWorks' built-in alignment.
        /// Ported from VBA AlignDims.bas::Align() (lines 10-52).
        /// </summary>
        public void AlignAllDimensions(IDrawingDoc drawDoc)
        {
            const string proc = nameof(AlignAllDimensions);
            try
            {
                var swModel = drawDoc as IModelDoc2;
                if (swModel == null) return;

                var sheet = drawDoc.GetCurrentSheet() as ISheet;
                if (sheet == null) return;

                var viewsRaw = sheet.GetViews() as object[];
                if (viewsRaw == null || viewsRaw.Length == 0) return;

                foreach (var viewObj in viewsRaw)
                {
                    var view = viewObj as IView;
                    if (view == null) continue;

                    var annots = view.GetAnnotations() as object[];
                    if (annots == null || annots.Length == 0) continue;

                    foreach (var annotObj in annots)
                    {
                        var annot = annotObj as IAnnotation;
                        if (annot == null) continue;
                        annot.Select3(true, null);
                    }

                    swModel.Extension.AlignDimensions(
                        (int)swAlignDimensionType_e.swAlignDimensionType_AutoArrange, 0.001);
                }

                swModel.ClearSelection2(true);
                swModel.GraphicsRedraw2();
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"{proc}: {ex.Message}");
            }
        }
    }
}
