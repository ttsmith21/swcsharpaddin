using System;
using System.Collections.Generic;
using System.Linq;
using NM.Core;
using NM.Core.Processing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Geometry
{
    /// <summary>
    /// Extracts tube/profile geometry from SolidWorks parts.
    /// Ported from VB.NET ExtractData CStepFile and CFaceCollection classes.
    /// </summary>
    public sealed class TubeGeometryExtractor
    {
        private const double Tolerance = 1e-9;

        private readonly ISldWorks _swApp;

        public TubeGeometryExtractor(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Extracts tube geometry from the active document.
        /// </summary>
        public TubeProfile ExtractFromActiveDocument()
        {
            var model = _swApp.ActiveDoc as IModelDoc2;
            if (model == null)
            {
                return new TubeProfile { Success = false, Message = "No active document" };
            }
            return Extract(model);
        }

        /// <summary>
        /// Extracts tube geometry from a model.
        /// </summary>
        public TubeProfile Extract(IModelDoc2 model)
        {
            const string proc = nameof(Extract);
            ErrorHandler.PushCallStack(proc);

            var result = new TubeProfile();

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

                var partDoc = (IPartDoc)model;
                var bodiesObj = partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
                if (bodiesObj == null)
                {
                    result.Message = "No solid bodies found";
                    return result;
                }

                var bodies = ((object[])bodiesObj).Cast<IBody2>().ToList();
                if (bodies.Count == 0)
                {
                    result.Message = "No solid bodies found";
                    return result;
                }

                // Collect all faces from all bodies
                var faces = new List<FaceWrapper>();
                foreach (var body in bodies)
                {
                    var bodyFaces = body.GetFaces() as object[];
                    if (bodyFaces != null)
                    {
                        faces.AddRange(bodyFaces.Cast<IFace2>().Select(f => new FaceWrapper(f)));
                    }
                }

                if (faces.Count == 0)
                {
                    result.Message = "No faces found";
                    return result;
                }

                // Extract geometry
                ExtractGeometry(model, faces, result);

                result.Success = result.Shape != TubeShape.None;
                if (!result.Success)
                {
                    result.Message = "Could not determine tube profile shape";
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

        private void ExtractGeometry(IModelDoc2 model, List<FaceWrapper> faces, TubeProfile result)
        {
            // Find max-area faces
            double maxArea = faces.Max(f => f.Area);
            var maxAreaFaces = faces.Where(f => IsTendsToZero(f.Area - maxArea)).ToList();

            // Check if any max-area face is cylindrical (round tube)
            var roundFace = maxAreaFaces.FirstOrDefault(f => f.IsRound);
            if (roundFace != null)
            {
                ExtractRoundProfile(model, faces, roundFace, result);
            }
            else
            {
                ExtractNonRoundProfile(model, faces, result);
            }
        }

        private void ExtractRoundProfile(IModelDoc2 model, List<FaceWrapper> faces, FaceWrapper outerFace, TubeProfile result)
        {
            result.Shape = TubeShape.Round;
            double outerRadius = outerFace.Radius;
            result.OuterDiameterMeters = outerRadius * 2;
            result.CrossSection = (outerRadius * 2).ToString("G6");

            // Find inner cylindrical face (parallel axis, smaller radius)
            double wallThickness = 0;
            foreach (var face in faces)
            {
                if (face == outerFace || !face.IsRound) continue;
                if (!outerFace.IsAxisParallelTo(face)) continue;

                double radiusDiff = Math.Abs(outerRadius - face.Radius);
                if (wallThickness == 0 || radiusDiff < wallThickness)
                {
                    wallThickness = radiusDiff;
                    result.InnerDiameterMeters = face.Radius * 2;
                }
            }
            result.WallThicknessMeters = wallThickness;

            // Get material length from outer cylindrical face
            double[] startPt, endPt;
            result.MaterialLengthMeters = outerFace.GetMaterialLength(out startPt, out endPt);
            result.StartPoint = startPt;
            result.EndPoint = endPt;

            // Count holes on the outer cylindrical face
            var holes = outerFace.GetHoles();
            result.NumberOfHoles = holes.Count;

            // Calculate cut length (edges at tube ends + hole edges)
            var cutEdges = new List<IEdge>();
            cutEdges.AddRange(outerFace.GetOuterLoopEdges());
            cutEdges.AddRange(outerFace.GetHoleEdges());
            result.CutLengthMeters = CalculateTotalEdgeLength(cutEdges);
        }

        private void ExtractNonRoundProfile(IModelDoc2 model, List<FaceWrapper> faces, TubeProfile result)
        {
            // Find largest planar face
            var planarFaces = faces.Where(f => f.IsPlanar).ToList();
            if (planarFaces.Count == 0)
            {
                result.Message = "No planar faces found";
                return;
            }

            double maxPlanarArea = planarFaces.Max(f => f.Area);
            var primaryFace = planarFaces.First(f => IsTendsToZero(f.Area - maxPlanarArea));

            // Get axis direction from largest linear edge
            double[] startPt, endPt;
            double edgeLength;
            var axisDirection = primaryFace.GetLargestLinearEdgeDirection(out startPt, out endPt, out edgeLength);
            result.StartPoint = startPt;
            result.EndPoint = endPt;
            result.MaterialLengthMeters = edgeLength;

            // Get cross-section dimensions by measuring parallel face distances
            var selMgr = (ISelectionMgr)model.SelectionManager;
            var selectData = selMgr.CreateSelectData();

            // Find faces parallel to primary face normal
            var primaryNormal = primaryFace.Normal;
            var facesParallelToPrimary = faces.Where(f => f.IsPlanar && f.IsNormalParallelTo(primaryFace)).ToList();

            // Find faces perpendicular to axis direction
            var facesNormalToAxis = faces.Where(f => f.IsPlanar && f.IsNormalParallelTo(axisDirection)).ToList();

            // Calculate cross product for third direction
            var secondaryDir = CrossProduct(primaryNormal, axisDirection);
            var facesNormalToSecondary = faces.Where(f => f.IsPlanar && f.IsNormalParallelTo(secondaryDir)).ToList();

            // Measure distances to determine cross-section
            double height = MeasureMaxDistance(model, selectData, facesParallelToPrimary, out var heightWall);
            double width = MeasureMaxDistance(model, selectData, facesNormalToSecondary, out var widthWall);
            double length = MeasureMaxDistance(model, selectData, facesNormalToAxis, out _);

            // Update material length if measured length is greater
            if (length > result.MaterialLengthMeters)
            {
                result.MaterialLengthMeters = length;
            }

            // Determine wall thickness (minimum of measured values)
            if (heightWall > 0 && widthWall > 0)
            {
                result.WallThicknessMeters = Math.Min(heightWall, widthWall);
            }
            else if (heightWall > 0)
            {
                result.WallThicknessMeters = heightWall;
            }
            else if (widthWall > 0)
            {
                result.WallThicknessMeters = widthWall;
            }

            // Determine shape
            bool heightZero = IsTendsToZero(height);
            bool widthZero = IsTendsToZero(width);

            if (heightZero && widthZero)
            {
                result.Shape = TubeShape.Angle;
            }
            else if (heightZero || widthZero)
            {
                result.Shape = TubeShape.Channel;
            }
            else if (IsTendsToZero(height - width))
            {
                result.Shape = TubeShape.Square;
            }
            else
            {
                result.Shape = TubeShape.Rectangle;
            }

            // Cross-section string
            double dim1 = Math.Max(height, width);
            double dim2 = Math.Min(height, width);
            if (dim1 > 0 && dim2 > 0)
            {
                result.CrossSection = $"{dim1:G6} x {dim2:G6}";
            }
            else if (dim1 > 0)
            {
                result.CrossSection = $"{dim1:G6}";
            }

            // Count holes and calculate cut length
            var cutFaces = new List<FaceWrapper>();
            cutFaces.AddRange(facesParallelToPrimary);
            cutFaces.AddRange(facesNormalToSecondary);

            int holeCount = 0;
            var cutEdges = new List<IEdge>();

            foreach (var face in cutFaces)
            {
                holeCount += face.GetHoles().Count;
                cutEdges.AddRange(face.GetHoleEdges());

                // Add outer edges that are perpendicular to axis (cut profile edges)
                foreach (var edge in face.GetOuterLoopEdges())
                {
                    if (!IsEdgeParallelToDirection(edge, axisDirection))
                    {
                        if (!cutEdges.Contains(edge))
                            cutEdges.Add(edge);
                    }
                }
            }

            result.NumberOfHoles = holeCount;
            result.CutLengthMeters = CalculateTotalEdgeLength(cutEdges);
        }

        private double MeasureMaxDistance(IModelDoc2 model, SelectData selectData, List<FaceWrapper> parallelFaces, out double wallThickness)
        {
            wallThickness = -1;
            double maxDistance = 0;

            if (parallelFaces.Count < 2)
                return 0;

            var measure = model.Extension.CreateMeasure();
            if (measure == null)
                return 0;

            measure.ArcOption = 0;

            for (int i = 0; i < parallelFaces.Count - 1; i++)
            {
                for (int j = i + 1; j < parallelFaces.Count; j++)
                {
                    model.ClearSelection2(true);
                    parallelFaces[i].SelectFace(selectData, false);
                    parallelFaces[j].SelectFace(selectData, true);

                    if (measure.Calculate(null))
                    {
                        double dist = measure.NormalDistance;
                        if (dist > 0)
                        {
                            if (wallThickness < 0)
                                wallThickness = dist;
                            else
                                wallThickness = Math.Min(wallThickness, dist);

                            maxDistance = Math.Max(maxDistance, dist);
                        }
                    }
                }
            }

            model.ClearSelection2(true);

            if (wallThickness < 0)
                wallThickness = 0;

            return maxDistance;
        }

        private double CalculateTotalEdgeLength(List<IEdge> edges)
        {
            double totalLength = 0;

            foreach (var edge in edges)
            {
                var curve = (ICurve)edge.GetCurve();
                if (curve == null) continue;

                double start = 0, end = 0;
                bool isClosed = false, isPeriodic = false;
                if (curve.GetEndParams(out start, out end, out isClosed, out isPeriodic))
                {
                    totalLength += curve.GetLength3(start, end);
                }
            }

            return totalLength;
        }

        private bool IsEdgeParallelToDirection(IEdge edge, double[] direction)
        {
            if (direction == null) return false;

            var curve = (ICurve)edge.GetCurve();
            if (curve == null || !curve.IsLine()) return false;

            var lineParams = (double[])curve.LineParams;
            if (lineParams == null || lineParams.Length < 6) return false;

            double[] edgeDir = { lineParams[3], lineParams[4], lineParams[5] };

            // Check if parallel (same or opposite direction)
            bool sameDir = IsTendsToZero(edgeDir[0] - direction[0]) &&
                           IsTendsToZero(edgeDir[1] - direction[1]) &&
                           IsTendsToZero(edgeDir[2] - direction[2]);
            bool oppDir = IsTendsToZero(edgeDir[0] + direction[0]) &&
                          IsTendsToZero(edgeDir[1] + direction[1]) &&
                          IsTendsToZero(edgeDir[2] + direction[2]);

            return sameDir || oppDir;
        }

        private static double[] CrossProduct(double[] v1, double[] v2)
        {
            return new double[]
            {
                v1[1] * v2[2] - v1[2] * v2[1],
                v1[2] * v2[0] - v1[0] * v2[2],
                v1[0] * v2[1] - v1[1] * v2[0]
            };
        }

        private static bool IsTendsToZero(double val)
        {
            return Math.Abs(val) < Tolerance;
        }
    }
}
