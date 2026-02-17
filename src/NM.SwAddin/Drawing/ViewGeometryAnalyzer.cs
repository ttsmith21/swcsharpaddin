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

        public ViewGeometryAnalyzer(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Finds and classifies all bend lines in a flat pattern drawing view.
        /// Separates into horizontal (angle=0) and vertical (angle=90) bends, sorted by position.
        /// </summary>
        /// <param name="drawDoc">The drawing document.</param>
        /// <param name="view">The flat pattern view.</param>
        /// <param name="horizontalBends">Output: horizontal bend lines sorted by Y position.</param>
        /// <param name="verticalBends">Output: vertical bend lines sorted by X position.</param>
        public void FindBendLines(IDrawingDoc drawDoc, IView view,
            out List<BendElement> horizontalBends, out List<BendElement> verticalBends)
        {
            // TODO: Phase 2 implementation
            // 1. Call view.GetBendLines() to get bend line sketch segments
            // 2. For each bend line, call ProcessLine() to get view-space coordinates
            // 3. Classify as horizontal (angle=0) or vertical (angle=90)
            // 4. Sort each list by Position, then P2
            horizontalBends = new List<BendElement>();
            verticalBends = new List<BendElement>();
        }

        /// <summary>
        /// Finds the boundary edge/vertex at the specified extreme of the view.
        /// Uses GetPolylines7 to traverse visible edges and find the most extreme position.
        /// </summary>
        /// <param name="view">The drawing view.</param>
        /// <param name="drawDoc">The drawing document.</param>
        /// <param name="direction">Which boundary to find (Left, Right, Top, Bottom).</param>
        /// <returns>The edge element at the boundary, or null if not found.</returns>
        public EdgeElement FindBoundaryEdge(IView view, IDrawingDoc drawDoc, BoundaryDirection direction)
        {
            // TODO: Phase 2 implementation
            // 1. Get edges via view.GetPolylines7(1, null)
            // 2. Filter to edges/silhouette edges with valid curves (not circles/ellipses/bcurves)
            // 3. ProcessLine() each to get view-space coordinates
            // 4. Find the extreme position for the requested direction
            // 5. Prefer full edges (Type="Line") over vertices (Type="Point")
            return null;
        }

        /// <summary>
        /// Transforms a line/edge/sketch segment into view-space coordinates and computes its angle.
        /// Handles bend lines (sketch segments), edges, and silhouette edges.
        /// </summary>
        /// <param name="line">The SolidWorks line object.</param>
        /// <param name="view">The drawing view.</param>
        /// <param name="drawDoc">The drawing document.</param>
        /// <param name="isBendLine">True if this is a bend line sketch segment.</param>
        /// <param name="startVertex">Output: start vertex (for Edge objects).</param>
        /// <param name="endVertex">Output: end vertex (for Edge objects).</param>
        /// <returns>The processed element with view-space coordinates and angle.</returns>
        public ProcessedElement ProcessLine(object line, IView view, IDrawingDoc drawDoc,
            bool isBendLine, out object startVertex, out object endVertex)
        {
            // TODO: Phase 2 implementation
            // 1. Determine object type (SketchSegment, Edge, SilhouetteEdge)
            // 2. Get start/end points using appropriate method:
            //    - BendLine: GetStartPoint2/GetEndPoint2 → TransformSketchPoint → ModelToViewTransform
            //    - Edge: GetStartVertex/GetEndVertex → CreatePoint → ModelToViewTransform
            //    - SilhouetteEdge: GetStartPoint/GetEndPoint → ModelToViewTransform
            // 3. Convert to view-space inches: coord * 39.3700787401575 / viewScale
            // 4. Compute angle: 0 if Y-delta=0, 90 if X-delta=0, else atan2
            startVertex = null;
            endVertex = null;
            return new ProcessedElement();
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
    }
}
