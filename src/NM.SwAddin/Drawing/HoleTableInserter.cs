using System;
using System.Collections.Generic;
using System.Linq;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Drawing
{
    /// <summary>
    /// Automatically inserts a hole table on flat pattern drawing views.
    /// Identifies datum vertex, holes face, and X/Y axis edges, then calls InsertHoleTable2.
    /// </summary>
    public sealed class HoleTableInserter
    {
        private readonly ISldWorks _swApp;

        /// <summary>
        /// Gap between view bottom edge and hole table top (meters).
        /// </summary>
        private const double TableGap = 0.005; // ~0.2"

        public HoleTableInserter(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Result of hole table insertion.
        /// </summary>
        public sealed class HoleTableResult
        {
            public bool Success { get; set; }
            public int HolesFound { get; set; }
            public string Message { get; set; }
        }

        /// <summary>
        /// Inserts a hole table on the given flat pattern view.
        /// Automatically identifies datum vertex, face, and X/Y axes.
        /// </summary>
        public HoleTableResult InsertHoleTable(IDrawingDoc drawDoc, IView flatPatternView, IModelDoc2 partModel)
        {
            const string proc = nameof(InsertHoleTable);
            ErrorHandler.PushCallStack(proc);
            PerformanceTracker.Instance.StartTimer("HoleTable_Insert");

            var result = new HoleTableResult();

            try
            {
                if (drawDoc == null || flatPatternView == null)
                {
                    result.Message = "DrawDoc or view is null";
                    return result;
                }

                var drawModel = drawDoc as IModelDoc2;
                string viewName = flatPatternView.GetName2();
                drawDoc.ActivateView(viewName);

                // Get visible components in the view to find entities
                var compObj = flatPatternView.GetVisibleComponents() as object[];
                IComponent2 viewComp = null;
                if (compObj != null && compObj.Length > 0)
                    viewComp = compObj[0] as IComponent2;

                // Get visible faces from the view
                var facesObj = flatPatternView.GetVisibleEntities2(
                    viewComp,
                    (int)swViewEntityType_e.swViewEntityType_Face) as object[];

                if (facesObj == null || facesObj.Length == 0)
                {
                    result.Message = "No visible faces in flat pattern view";
                    return result;
                }

                // Find the largest planar face (flat pattern main face)
                IFace2 largestFace = null;
                double largestArea = 0;

                foreach (var fObj in facesObj)
                {
                    var face = fObj as IFace2;
                    if (face == null) continue;

                    var surface = face.GetSurface() as ISurface;
                    if (surface == null || !surface.IsPlane()) continue;

                    double area = face.GetArea();
                    if (area > largestArea)
                    {
                        largestArea = area;
                        largestFace = face;
                    }
                }

                if (largestFace == null)
                {
                    result.Message = "No planar faces found in flat pattern view";
                    return result;
                }

                // Get outer loop edges of the largest face
                var outerEdges = GetOuterLoopEdges(largestFace);
                if (outerEdges.Count == 0)
                {
                    result.Message = "No outer loop edges found on flat pattern face";
                    return result;
                }

                // Find bottom-left vertex as datum origin
                IVertex datumVertex = null;
                double minSum = double.MaxValue;

                foreach (var edge in outerEdges)
                {
                    foreach (var vertex in new[] { edge.GetStartVertex() as IVertex, edge.GetEndVertex() as IVertex })
                    {
                        if (vertex == null) continue;
                        var pt = vertex.GetPoint() as double[];
                        if (pt == null || pt.Length < 2) continue;

                        double sum = pt[0] + pt[1]; // X + Y — minimum = bottom-left
                        if (sum < minSum)
                        {
                            minSum = sum;
                            datumVertex = vertex;
                        }
                    }
                }

                if (datumVertex == null)
                {
                    result.Message = "Could not find datum vertex";
                    return result;
                }

                // Find X-axis (horizontal) and Y-axis (vertical) edges touching the datum vertex
                var datumPt = datumVertex.GetPoint() as double[];
                IEdge xAxisEdge = null;
                IEdge yAxisEdge = null;
                double bestHorizLength = 0;
                double bestVertLength = 0;

                foreach (var edge in outerEdges)
                {
                    var curve = edge.GetCurve() as ICurve;
                    if (curve == null || !curve.IsLine()) continue;

                    // Check if this edge touches the datum vertex
                    if (!EdgeTouchesVertex(edge, datumPt)) continue;

                    var startV = edge.GetStartVertex() as IVertex;
                    var endV = edge.GetEndVertex() as IVertex;
                    if (startV == null || endV == null) continue;

                    var startPt = startV.GetPoint() as double[];
                    var endPt = endV.GetPoint() as double[];
                    if (startPt == null || endPt == null) continue;

                    double dx = Math.Abs(endPt[0] - startPt[0]);
                    double dy = Math.Abs(endPt[1] - startPt[1]);
                    double length = Math.Sqrt(dx * dx + dy * dy);

                    // Classify as horizontal or vertical
                    if (dx > dy && length > bestHorizLength)
                    {
                        bestHorizLength = length;
                        xAxisEdge = edge;
                    }
                    else if (dy > dx && length > bestVertLength)
                    {
                        bestVertLength = length;
                        yAxisEdge = edge;
                    }
                }

                // If no edges touch datum vertex exactly, find closest horizontal/vertical outer edges
                if (xAxisEdge == null || yAxisEdge == null)
                {
                    foreach (var edge in outerEdges)
                    {
                        var curve = edge.GetCurve() as ICurve;
                        if (curve == null || !curve.IsLine()) continue;

                        var startV = edge.GetStartVertex() as IVertex;
                        var endV = edge.GetEndVertex() as IVertex;
                        if (startV == null || endV == null) continue;

                        var startPt = startV.GetPoint() as double[];
                        var endPt = endV.GetPoint() as double[];
                        if (startPt == null || endPt == null) continue;

                        double dx = Math.Abs(endPt[0] - startPt[0]);
                        double dy = Math.Abs(endPt[1] - startPt[1]);
                        double length = Math.Sqrt(dx * dx + dy * dy);

                        if (xAxisEdge == null && dx > dy && length > bestHorizLength)
                        {
                            bestHorizLength = length;
                            xAxisEdge = edge;
                        }
                        if (yAxisEdge == null && dy > dx && length > bestVertLength)
                        {
                            bestVertLength = length;
                            yAxisEdge = edge;
                        }
                    }
                }

                if (xAxisEdge == null || yAxisEdge == null)
                {
                    result.Message = "Could not identify X/Y axis edges for hole table";
                    return result;
                }

                // Pre-select entities with marks
                drawModel.ClearSelection2(true);
                var selMgr = drawModel.SelectionManager as ISelectionMgr;
                var selData = selMgr.CreateSelectData() as ISelectData;

                // Mark 1: datum vertex (origin)
                selData.Mark = 1;
                ((IEntity)datumVertex).Select4(false, selData);

                // Mark 2: holes face
                selData.Mark = 2;
                ((IEntity)largestFace).Select4(true, selData);

                // Mark 4: X-axis edge
                selData.Mark = 4;
                ((IEntity)xAxisEdge).Select4(true, selData);

                // Mark 8: Y-axis edge
                selData.Mark = 8;
                ((IEntity)yAxisEdge).Select4(true, selData);

                // Calculate position below the view
                var viewOutline = flatPatternView.GetOutline() as double[];
                double xPos = 0.01; // Default
                double yPos = 0.01;
                if (viewOutline != null && viewOutline.Length >= 4)
                {
                    xPos = viewOutline[0]; // Left-aligned with view
                    yPos = viewOutline[1] - TableGap; // Below the view
                }

                // Insert the hole table
                ErrorHandler.DebugLog($"[HoleTable] Inserting hole table on '{viewName}' at ({xPos:F4}, {yPos:F4})");

                var holeTable = flatPatternView.InsertHoleTable2(
                    false,    // UseAnchorPoint
                    xPos,     // X position (meters)
                    yPos,     // Y position (meters)
                    (int)swBomBalloonStyle_e.swBS_Circular,
                    "",       // Template (empty = default)
                    ""        // Origin indicator
                ) as IHoleTableAnnotation;

                if (holeTable == null)
                {
                    result.Message = "InsertHoleTable2 returned null (no holes detected or selection error)";
                    // Not necessarily a failure — part may have no holes
                    result.Success = true;
                    result.HolesFound = 0;
                    return result;
                }

                // Count holes in the table
                var tableAnn = holeTable as ITableAnnotation;
                if (tableAnn != null)
                {
                    result.HolesFound = Math.Max(0, tableAnn.RowCount - 1); // Subtract header row
                }

                result.Success = true;
                result.Message = $"Hole table inserted with {result.HolesFound} holes";
                ErrorHandler.DebugLog($"[HoleTable] {result.Message}");

                return result;
            }
            catch (Exception ex)
            {
                result.Message = "Exception: " + ex.Message;
                ErrorHandler.DebugLog($"[HoleTable] Error: {ex.Message}");
                return result;
            }
            finally
            {
                PerformanceTracker.Instance.StopTimer("HoleTable_Insert");
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Gets outer loop edges from a face.
        /// </summary>
        private List<IEdge> GetOuterLoopEdges(IFace2 face)
        {
            var edges = new List<IEdge>();

            var loopsRaw = face.GetLoops() as object[];
            if (loopsRaw == null) return edges;

            foreach (var loopObj in loopsRaw)
            {
                var loop = loopObj as ILoop2;
                if (loop == null || !loop.IsOuter()) continue;

                var loopEdges = loop.GetEdges() as object[];
                if (loopEdges == null) continue;

                foreach (var edgeObj in loopEdges)
                {
                    var edge = edgeObj as IEdge;
                    if (edge != null)
                        edges.Add(edge);
                }
            }

            return edges;
        }

        /// <summary>
        /// Checks if an edge has an endpoint at the given position (within tolerance).
        /// </summary>
        private bool EdgeTouchesVertex(IEdge edge, double[] vertexPt)
        {
            const double tol = 1e-6;

            var startV = edge.GetStartVertex() as IVertex;
            var endV = edge.GetEndVertex() as IVertex;

            if (startV != null)
            {
                var pt = startV.GetPoint() as double[];
                if (pt != null && Math.Abs(pt[0] - vertexPt[0]) < tol &&
                    Math.Abs(pt[1] - vertexPt[1]) < tol)
                    return true;
            }

            if (endV != null)
            {
                var pt = endV.GetPoint() as double[];
                if (pt != null && Math.Abs(pt[0] - vertexPt[0]) < tol &&
                    Math.Abs(pt[1] - vertexPt[1]) < tol)
                    return true;
            }

            return false;
        }
    }
}
