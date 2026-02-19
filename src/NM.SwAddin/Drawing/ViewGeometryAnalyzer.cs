using System;
using System.Collections.Generic;
using System.Linq;
using NM.Core;
using NM.Core.Drawing;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.Drawing
{
    /// <summary>
    /// Analyzes drawing view geometry: finds bend lines, boundary edges, and transforms
    /// coordinates between model/sketch/view space.
    /// Ported from VBA DimensionDrawing.bas: FindBendLines, FindLeftPosLine,
    /// FindRightPosLine, FindTopPosLine, FindBottomPosLine, ProcessLine,
    /// TransformSketchPointToModelSpace.
    /// </summary>
    public sealed class ViewGeometryAnalyzer
    {
        private readonly ISldWorks _swApp;

        /// <summary>Inches per meter, same constant used in VBA.</summary>
        private const double InchesPerMeter = 39.3700787401575;

        public ViewGeometryAnalyzer(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Finds and classifies all bend lines in a flat pattern drawing view.
        /// Separates into horizontal (angle=0) and vertical (angle=90) bends, sorted by position.
        /// Ported from VBA FindBendLines() (lines 677-723).
        /// </summary>
        public void FindBendLines(IDrawingDoc drawDoc, IView view,
            out List<BendElement> horizontalBends, out List<BendElement> verticalBends)
        {
            horizontalBends = new List<BendElement>();
            verticalBends = new List<BendElement>();

            var vBendLines = view.GetBendLines() as object[];
            if (vBendLines == null || vBendLines.Length == 0) return;

            for (int i = 0; i < vBendLines.Length; i++)
            {
                var bendLine = vBendLines[i];
                if (bendLine == null) continue;

                object startVertex = null;
                object endVertex = null;
                var processed = ProcessLine(bendLine, view, drawDoc, true, out startVertex, out endVertex);
                if (processed == null) continue;

                if (processed.Angle == 0)
                {
                    // Horizontal bend: Position = Y1, P2 = min(X1,X2)
                    horizontalBends.Add(new BendElement
                    {
                        SketchSegment = processed.Obj,
                        Position = processed.Y1,
                        P2 = Math.Min(processed.X1, processed.X2),
                        Angle = processed.Angle,
                    });
                }
                else if (processed.Angle == 90)
                {
                    // Vertical bend: Position = X1, P2 = min(Y1,Y2)
                    verticalBends.Add(new BendElement
                    {
                        SketchSegment = processed.Obj,
                        Position = processed.X1,
                        P2 = Math.Min(processed.Y1, processed.Y2),
                        Angle = processed.Angle,
                    });
                }
            }

            SortBendElements(horizontalBends);
            SortBendElements(verticalBends);
        }

        /// <summary>
        /// Finds the boundary edge/vertex at the specified extreme of the view.
        /// Uses GetPolylines7 to traverse visible edges and find the most extreme position.
        /// Ported from VBA FindLeftPosLine/FindRightPosLine/FindTopPosLine/FindBottomPosLine (lines 724-948).
        /// </summary>
        public EdgeElement FindBoundaryEdge(IView view, IDrawingDoc drawDoc, BoundaryDirection direction)
        {
            var result = new EdgeElement();

            // Initialize extreme values
            switch (direction)
            {
                case BoundaryDirection.Left:   result.X = 9999; break;
                case BoundaryDirection.Right:  result.X = 0; break;
                case BoundaryDirection.Top:    result.Y = 0; break;
                case BoundaryDirection.Bottom: result.Y = 9999; break;
            }

            object polylineData;
            var vViewLines = view.GetPolylines7(1, out polylineData) as object[];
            if (vViewLines == null || vViewLines.Length == 0) return null;

            for (int i = 0; i < vViewLines.Length; i++)
            {
                var swLine = vViewLines[i];
                if (swLine == null) continue;

                // Filter: must be Edge or SilhouetteEdge
                bool isEdge = swLine is IEdge;
                bool isSilhouette = !isEdge && (swLine is ISilhouetteEdge);
                if (!isEdge && !isSilhouette) continue;

                // Filter: must have a curve that is not circle/ellipse/bcurve
                ICurve curve = null;
                try
                {
                    if (isEdge)
                        curve = ((IEdge)swLine).GetCurve() as ICurve;
                    else
                        curve = ((ISilhouetteEdge)swLine).GetCurve() as ICurve;
                }
                catch { continue; }

                if (curve == null) continue;
                // Filter out non-line curves: Identity 3001 = LINE_TYPE
                // VBA checks: IsCircle, IsEllipse, IsBcurve — we use Identity for SW 2022 compat
                int curveType = curve.Identity();
                if (curveType != 3001) continue; // Only keep lines

                object startVertex = null;
                object endVertex = null;
                var processed = ProcessLine(swLine, view, drawDoc, false, out startVertex, out endVertex);
                if (processed == null) continue;

                switch (direction)
                {
                    case BoundaryDirection.Left:
                        UpdateLeftBoundary(ref result, processed, startVertex, endVertex);
                        break;
                    case BoundaryDirection.Right:
                        UpdateRightBoundary(ref result, processed, startVertex, endVertex);
                        break;
                    case BoundaryDirection.Top:
                        UpdateTopBoundary(ref result, processed, startVertex, endVertex);
                        break;
                    case BoundaryDirection.Bottom:
                        UpdateBottomBoundary(ref result, processed, startVertex, endVertex);
                        break;
                }
            }

            // Return null if no edge was found (obj still null)
            return result.Obj != null ? result : null;
        }

        /// <summary>
        /// Transforms a line/edge/sketch segment into view-space coordinates and computes its angle.
        /// Ported from VBA ProcessLine() (lines 1086-1157).
        /// </summary>
        public ProcessedElement ProcessLine(object line, IView view, IDrawingDoc drawDoc,
            bool isBendLine, out object startVertex, out object endVertex)
        {
            startVertex = null;
            endVertex = null;

            try
            {
                // Get view scale from GetXform: xform[2] = scale
                var xform = view.GetXform() as double[];
                if (xform == null || xform.Length < 3) return null;
                double viewScale = xform[2];
                if (viewScale == 0) viewScale = 1;

                var mathUtil = _swApp.GetMathUtility() as IMathUtility;
                if (mathUtil == null) return null;

                IMathPoint swViewStartPt = null;
                IMathPoint swViewEndPt = null;
                var viewXform = view.ModelToViewTransform;

                if (isBendLine)
                {
                    // Bend line: ISketchLine → GetStartPoint2/GetEndPoint2
                    var sketchLine = line as ISketchLine;
                    if (sketchLine == null) return null;

                    var sketchSeg = line as ISketchSegment;
                    var sketch = sketchSeg?.GetSketch() as ISketch;
                    if (sketch == null) return null;

                    var pStart = sketchLine.GetStartPoint2() as ISketchPoint;
                    var pEnd = sketchLine.GetEndPoint2() as ISketchPoint;
                    if (pStart == null || pEnd == null) return null;

                    var modelStartPt = TransformSketchPointToModelSpace(mathUtil, sketch, pStart);
                    var modelEndPt = TransformSketchPointToModelSpace(mathUtil, sketch, pEnd);
                    if (modelStartPt == null || modelEndPt == null) return null;

                    swViewStartPt = modelStartPt.MultiplyTransform(viewXform) as IMathPoint;
                    swViewEndPt = modelEndPt.MultiplyTransform(viewXform) as IMathPoint;
                }
                else if (line is ISilhouetteEdge)
                {
                    var silEdge = (ISilhouetteEdge)line;
                    var modelStartPt = silEdge.GetStartPoint() as IMathPoint;
                    var modelEndPt = silEdge.GetEndPoint() as IMathPoint;
                    if (modelStartPt == null || modelEndPt == null) return null;

                    swViewStartPt = modelStartPt.MultiplyTransform(viewXform) as IMathPoint;
                    swViewEndPt = modelEndPt.MultiplyTransform(viewXform) as IMathPoint;
                }
                else if (line is IEdge)
                {
                    var edge = (IEdge)line;
                    var sv = edge.GetStartVertex() as IVertex;
                    var ev = edge.GetEndVertex() as IVertex;
                    if (sv == null || ev == null) return null;

                    startVertex = sv;
                    endVertex = ev;

                    var startPtArr = sv.GetPoint() as double[];
                    var endPtArr = ev.GetPoint() as double[];
                    if (startPtArr == null || endPtArr == null) return null;

                    var modelStartPt = mathUtil.CreatePoint(startPtArr) as IMathPoint;
                    var modelEndPt = mathUtil.CreatePoint(endPtArr) as IMathPoint;
                    if (modelStartPt == null || modelEndPt == null) return null;

                    swViewStartPt = modelStartPt.MultiplyTransform(viewXform) as IMathPoint;
                    swViewEndPt = modelEndPt.MultiplyTransform(viewXform) as IMathPoint;
                }
                else
                {
                    // Other sketch segment type — try ISketchLine first
                    var sketchLine = line as ISketchLine;
                    if (sketchLine == null) return null;

                    var sketchSeg = line as ISketchSegment;
                    var sketch = sketchSeg?.GetSketch() as ISketch;
                    if (sketch == null) return null;

                    var pStart = sketchLine.GetStartPoint2() as ISketchPoint;
                    var pEnd = sketchLine.GetEndPoint2() as ISketchPoint;
                    if (pStart == null || pEnd == null) return null;

                    var modelStartPt = TransformSketchPointToModelSpace(mathUtil, sketch, pStart);
                    var modelEndPt = TransformSketchPointToModelSpace(mathUtil, sketch, pEnd);
                    if (modelStartPt == null || modelEndPt == null) return null;

                    swViewStartPt = modelStartPt.MultiplyTransform(viewXform) as IMathPoint;
                    swViewEndPt = modelEndPt.MultiplyTransform(viewXform) as IMathPoint;
                }

                if (swViewStartPt == null || swViewEndPt == null) return null;

                var startArr = swViewStartPt.ArrayData as double[];
                var endArr = swViewEndPt.ArrayData as double[];
                if (startArr == null || endArr == null) return null;

                // Convert to view-space inches: coord * InchesPerMeter / viewScale
                double x1 = Math.Round(startArr[0] * InchesPerMeter / viewScale, 3);
                double x2 = Math.Round(endArr[0] * InchesPerMeter / viewScale, 3);
                double y1 = Math.Round(startArr[1] * InchesPerMeter / viewScale, 3);
                double y2 = Math.Round(endArr[1] * InchesPerMeter / viewScale, 3);

                // Compute angle
                double angle;
                if (y2 - y1 == 0)
                    angle = 0;
                else if (x2 - x1 == 0)
                    angle = 90;
                else
                    angle = Math.Round(Math.Atan2(y2 - y1, x2 - x1) * 180.0 / Math.PI, 3);

                return new ProcessedElement
                {
                    Obj = line,
                    X1 = x1,
                    X2 = x2,
                    Y1 = y1,
                    Y2 = y2,
                    Angle = angle,
                };
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"ProcessLine: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Transforms a sketch point from sketch space to model space.
        /// Ported from VBA TransformSketchPointToModelSpace() (lines 1158-1190).
        /// </summary>
        private IMathPoint TransformSketchPointToModelSpace(IMathUtility mathUtil, ISketch sketch, ISketchPoint skPt)
        {
            try
            {
                var nPt = new double[] { skPt.X, skPt.Y, skPt.Z };
                var sketchXform = sketch.ModelToSketchTransform;
                var inverseXform = sketchXform.Inverse();
                var mathPt = mathUtil.CreatePoint(nPt) as IMathPoint;
                return mathPt?.MultiplyTransform(inverseXform) as IMathPoint;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Sorts bend elements by Position, then by P2 for co-located bends.
        /// </summary>
        public static void SortBendElements(List<BendElement> bends)
        {
            bends.Sort((a, b) =>
            {
                int cmp = a.Position.CompareTo(b.Position);
                return cmp != 0 ? cmp : a.P2.CompareTo(b.P2);
            });
        }

        #region Boundary edge helpers

        /// <summary>VBA FindLeftPosLine logic: find edge with minimum X.</summary>
        private void UpdateLeftBoundary(ref EdgeElement result, ProcessedElement processed,
            object startVertex, object endVertex)
        {
            // Check X1 (start point)
            if (processed.X1 < result.X)
            {
                if (processed.Angle == 90)
                {
                    result.Obj = processed.Obj;
                    result.Type = "Line";
                    result.Y = Math.Min(processed.Y1, processed.Y2);
                }
                else
                {
                    result.Obj = startVertex ?? processed.Obj;
                    result.Type = "Point";
                    result.Y = processed.Y1;
                }
                result.X = processed.X1;
                result.Angle = processed.Angle;
            }
            else if (processed.X1 == result.X && processed.Angle == 90)
            {
                double minY = Math.Min(processed.Y1, processed.Y2);
                if (minY < result.Y || result.Type != "Line")
                {
                    result.Obj = processed.Obj;
                    result.Type = "Line";
                    result.X = processed.X1;
                    result.Y = minY;
                    result.Angle = processed.Angle;
                }
            }

            // Check X2 (end point)
            if (processed.X2 < result.X)
            {
                result.Obj = endVertex ?? processed.Obj;
                result.Type = "Point";
                result.X = processed.X2;
                result.Y = processed.Y2;
                result.Angle = processed.Angle;
            }
        }

        /// <summary>VBA FindRightPosLine logic: find edge with maximum X.</summary>
        private void UpdateRightBoundary(ref EdgeElement result, ProcessedElement processed,
            object startVertex, object endVertex)
        {
            if (processed.X1 > result.X)
            {
                if (processed.Angle == 90)
                {
                    result.Obj = processed.Obj;
                    result.Type = "Line";
                    result.Y = Math.Min(processed.Y1, processed.Y2);
                }
                else
                {
                    result.Obj = startVertex ?? processed.Obj;
                    result.Type = "Point";
                    result.Y = processed.Y1;
                }
                result.X = processed.X1;
                result.Angle = processed.Angle;
            }
            else if (processed.X1 == result.X && processed.Angle == 90)
            {
                double minY = Math.Min(processed.Y1, processed.Y2);
                if (minY < result.Y || result.Type != "Line")
                {
                    result.Obj = processed.Obj;
                    result.Type = "Line";
                    result.X = processed.X1;
                    result.Y = minY;
                    result.Angle = processed.Angle;
                }
            }

            if (processed.X2 > result.X)
            {
                result.Obj = endVertex ?? processed.Obj;
                result.Type = "Point";
                result.X = processed.X2;
                result.Y = processed.Y2;
                result.Angle = processed.Angle;
            }
        }

        /// <summary>VBA FindTopPosLine logic: find edge with maximum Y.</summary>
        private void UpdateTopBoundary(ref EdgeElement result, ProcessedElement processed,
            object startVertex, object endVertex)
        {
            if (processed.Y1 > result.Y)
            {
                if (processed.Angle == 0)
                {
                    result.Obj = processed.Obj;
                    result.Type = "Line";
                    result.X = Math.Min(processed.X1, processed.X2);
                }
                else
                {
                    result.Obj = startVertex ?? processed.Obj;
                    result.Type = "Point";
                    result.X = processed.X1;
                }
                result.Y = processed.Y1;
                result.Angle = processed.Angle;
            }
            else if (processed.Y1 == result.Y && processed.Angle == 0)
            {
                double minX = Math.Min(processed.X1, processed.X2);
                if (minX < result.X || result.Type != "Line")
                {
                    result.Obj = processed.Obj;
                    result.Type = "Line";
                    result.Y = processed.Y1;
                    result.X = minX;
                    result.Angle = processed.Angle;
                }
            }

            if (processed.Y2 > result.Y)
            {
                result.Obj = endVertex ?? processed.Obj;
                result.Type = "Point";
                result.Y = processed.Y2;
                result.X = processed.X2;
                result.Angle = processed.Angle;
            }
        }

        /// <summary>VBA FindBottomPosLine logic: find edge with minimum Y.</summary>
        private void UpdateBottomBoundary(ref EdgeElement result, ProcessedElement processed,
            object startVertex, object endVertex)
        {
            if (processed.Y1 < result.Y)
            {
                if (processed.Angle == 0)
                {
                    result.Obj = processed.Obj;
                    result.Type = "Line";
                    result.X = Math.Min(processed.X1, processed.X2);
                }
                else
                {
                    result.Obj = startVertex ?? processed.Obj;
                    result.Type = "Point";
                    result.X = processed.X1;
                }
                result.Y = processed.Y1;
                result.Angle = processed.Angle;
            }
            else if (processed.Y1 == result.Y && processed.Angle == 0)
            {
                double minX = Math.Min(processed.X1, processed.X2);
                if (minX < result.X || result.Type != "Line")
                {
                    result.Obj = processed.Obj;
                    result.Type = "Line";
                    result.Y = processed.Y1;
                    result.X = minX;
                    result.Angle = processed.Angle;
                }
            }

            if (processed.Y2 < result.Y)
            {
                result.Obj = endVertex ?? processed.Obj;
                result.Type = "Point";
                result.Y = processed.Y2;
                result.X = processed.X2;
                result.Angle = processed.Angle;
            }
        }

        #endregion
    }
}
