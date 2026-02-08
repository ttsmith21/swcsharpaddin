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

            // Get material length from outer loop edges (circles at cylinder ends)
            // This is more reliable than tessellation for cylindrical faces
            var outerEdges = outerFace.GetOuterLoopEdges();
            result.MaterialLengthMeters = CalculateCylinderLengthFromEdges(outerEdges, outerFace.Axis, out var startPt, out var endPt);
            result.StartPoint = startPt;
            result.EndPoint = endPt;

            // Count holes on the outer cylindrical face
            var holes = outerFace.GetHoles();
            result.NumberOfHoles = holes.Count;

            // Calculate cut length (edges at tube ends + hole edges)
            var cutEdges = new List<IEdge>();
            cutEdges.AddRange(outerEdges);
            cutEdges.AddRange(outerFace.GetHoleEdges());
            result.CutLengthMeters = CalculateTotalEdgeLength(cutEdges);
        }

        /// <summary>
        /// Calculates cylinder length by finding the distance between edge loop centers along the axis.
        /// </summary>
        private double CalculateCylinderLengthFromEdges(List<IEdge> edges, double[] axis, out double[] startPoint, out double[] endPoint)
        {
            startPoint = new double[3];
            endPoint = new double[3];

            if (edges == null || edges.Count < 2 || axis == null)
                return 0;

            // Project edge midpoints onto the axis to find min/max positions
            double minT = double.MaxValue;
            double maxT = double.MinValue;
            double[] minPt = null;
            double[] maxPt = null;

            // Use axis origin from the first edge curve's center
            double[] axisOrigin = null;

            foreach (var edge in edges)
            {
                var curve = (ICurve)edge.GetCurve();
                if (curve == null) continue;

                // Get curve midpoint
                double startParam = 0, endParam = 0;
                bool isClosed = false, isPeriodic = false;
                if (!curve.GetEndParams(out startParam, out endParam, out isClosed, out isPeriodic))
                    continue;

                double midParam = (startParam + endParam) / 2.0;
                int numDerivs = 0;
                var evalResult = (double[])curve.Evaluate2(midParam, numDerivs);
                if (evalResult == null || evalResult.Length < 3)
                    continue;

                double[] midPt = { evalResult[0], evalResult[1], evalResult[2] };

                // Initialize axis origin with first point
                if (axisOrigin == null)
                    axisOrigin = midPt;

                // Project onto axis: t = (midPt - axisOrigin) · axis
                double t = (midPt[0] - axisOrigin[0]) * axis[0] +
                           (midPt[1] - axisOrigin[1]) * axis[1] +
                           (midPt[2] - axisOrigin[2]) * axis[2];

                if (t < minT)
                {
                    minT = t;
                    minPt = midPt;
                }
                if (t > maxT)
                {
                    maxT = t;
                    maxPt = midPt;
                }
            }

            if (minPt != null && maxPt != null && maxT > minT)
            {
                startPoint = minPt;
                endPoint = maxPt;
                return maxT - minT;
            }

            return 0;
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

            // Get cross-section dimensions by computing parallel face distances geometrically
            // Find faces parallel to primary face normal
            var primaryNormal = primaryFace.Normal;
            var facesParallelToPrimary = faces.Where(f => f.IsPlanar && f.IsNormalParallelTo(primaryFace)).ToList();

            // Find faces perpendicular to axis direction
            var facesNormalToAxis = faces.Where(f => f.IsPlanar && f.IsNormalParallelTo(axisDirection)).ToList();

            // Calculate cross product for third direction
            var secondaryDir = CrossProduct(primaryNormal, axisDirection);
            var facesNormalToSecondary = faces.Where(f => f.IsPlanar && f.IsNormalParallelTo(secondaryDir)).ToList();

            // Compute cross-section dimensions using dot-product plane distances (no COM Measure calls)
            int distinctDistancesPrimary, distinctDistancesSecondary;
            double height = MeasureMaxDistanceWithValidation(facesParallelToPrimary, out var heightWall, out distinctDistancesPrimary);
            double width = MeasureMaxDistanceWithValidation(facesNormalToSecondary, out var widthWall, out distinctDistancesSecondary);
            double length = MeasureMaxDistance(facesNormalToAxis, out _);

            // Diagnostic logging
            ErrorHandler.DebugLog($"[TUBE] ExtractNonRoundProfile measurements: " +
                $"height={height * 39.37:F3}in, width={width * 39.37:F3}in, " +
                $"heightWall={heightWall * 39.37:F3}in, widthWall={widthWall * 39.37:F3}in, " +
                $"primaryFaces={facesParallelToPrimary.Count}, secondaryFaces={facesNormalToSecondary.Count}, " +
                $"distinctDist1={distinctDistancesPrimary}, distinctDist2={distinctDistancesSecondary}");

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

            // Shape determination — faithful port of VB.NET CFaceCollection.ComputeShape()
            // Tier 1: Zero-dimension check (open profiles where no parallel faces exist in that direction)
            if (IsTendsToZero(height) || IsTendsToZero(width))
            {
                if (IsTendsToZero(height) && IsTendsToZero(width))
                    result.Shape = TubeShape.Angle;
                else
                    result.Shape = TubeShape.Channel;
            }
            else
            {
                // Tier 2: Distance count modulo check (structural shapes with non-zero dimensions)
                // Odd count in a direction = open profile in that direction (e.g., angle leg, channel flange)
                int remPrimary = distinctDistancesPrimary % 2;
                int remSecondary = distinctDistancesSecondary % 2;

                if (remPrimary != 0 || remSecondary != 0)
                {
                    if (remPrimary != 0 && remSecondary != 0)
                        result.Shape = TubeShape.Angle;   // Both directions open
                    else
                        result.Shape = TubeShape.Channel;  // One direction open
                }
                else
                {
                    // Both directions have even distance counts = enclosed profile
                    if (IsTendsToZero(height - width))
                        result.Shape = TubeShape.Square;
                    else
                        result.Shape = TubeShape.Rectangle;
                }
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

            // Store cross-section dimensions for OptiMaterial generation
            // For non-round profiles: OD = larger dimension, ID = smaller dimension
            result.OuterDiameterMeters = dim1;
            result.InnerDiameterMeters = dim2;

            // Count holes and calculate cut length
            // VB.NET Audit Resolution: Filter to largest-area faces and use boundary edge detection
            var cutFaces = new List<FaceWrapper>();
            cutFaces.AddRange(facesParallelToPrimary);
            cutFaces.AddRange(facesNormalToSecondary);

            // Filter to largest-area faces only (VB.NET face removal logic)
            var primaryCutFaces = FilterToLargestAreaFaces(cutFaces);

            int holeCount = 0;
            var cutEdges = new List<IEdge>();

            // Collect hole edges from all primary faces
            foreach (var face in primaryCutFaces)
            {
                holeCount += face.GetHoles().Count;
                cutEdges.AddRange(face.GetHoleEdges());
            }

            // Get boundary edges using VB.NET audit resolution method
            // These are edges where one adjacent face is in our collection and one is not
            var boundaryEdges = FaceWrapper.GetBoundaryEdges(primaryCutFaces, _swApp);

            // Filter boundary edges: only keep those perpendicular to axis
            foreach (var edge in boundaryEdges)
            {
                if (!IsEdgeParallelToDirection(edge, axisDirection))
                {
                    if (!cutEdges.Contains(edge))
                        cutEdges.Add(edge);
                }
            }

            // Fallback: if boundary detection yields no edges, use original method
            if (boundaryEdges.Count == 0)
            {
                ErrorHandler.DebugLog("[TUBE] Boundary edge detection returned no edges, using fallback");
                foreach (var face in primaryCutFaces)
                {
                    foreach (var edge in face.GetOuterLoopEdges())
                    {
                        if (!IsEdgeParallelToDirection(edge, axisDirection))
                        {
                            if (!cutEdges.Contains(edge))
                                cutEdges.Add(edge);
                        }
                    }
                }
            }

            result.NumberOfHoles = holeCount;
            result.CutLengthMeters = CalculateTotalEdgeLength(cutEdges);
        }

        private double MeasureMaxDistance(List<FaceWrapper> parallelFaces, out double wallThickness)
        {
            int distinctDistances;
            return MeasureMaxDistanceWithValidation(parallelFaces, out wallThickness, out distinctDistances);
        }

        /// <summary>
        /// Computes distances between parallel planar faces using dot-product geometry.
        /// Returns the count of distinct distances found.
        /// Used for validation: true extrusions have paired faces (even count), machined parts have odd counts.
        /// </summary>
        private double MeasureMaxDistanceWithValidation(List<FaceWrapper> parallelFaces,
            out double wallThickness, out int distinctDistanceCount)
        {
            wallThickness = -1;
            distinctDistanceCount = 0;
            double maxDistance = 0;

            if (parallelFaces.Count < 2)
                return 0;

            // Track distinct distances (within tolerance) for the modulo check
            var distinctDistances = new List<double>();
            const double DISTANCE_TOLERANCE = 0.0001; // 0.1mm tolerance for grouping distances

            for (int i = 0; i < parallelFaces.Count - 1; i++)
            {
                var n = parallelFaces[i].Normal;
                var p1 = parallelFaces[i].PlaneOrigin;
                if (p1 == null) continue;

                for (int j = i + 1; j < parallelFaces.Count; j++)
                {
                    var p2 = parallelFaces[j].PlaneOrigin;
                    if (p2 == null) continue;

                    // Distance between parallel planes: |dot(normal, p2 - p1)|
                    double dist = Math.Abs(
                        n[0] * (p2[0] - p1[0]) +
                        n[1] * (p2[1] - p1[1]) +
                        n[2] * (p2[2] - p1[2])
                    );

                    if (dist > 1e-6) // Filter floating-point noise for coplanar faces
                    {
                        if (wallThickness < 0)
                            wallThickness = dist;
                        else
                            wallThickness = Math.Min(wallThickness, dist);

                        maxDistance = Math.Max(maxDistance, dist);

                        // Track distinct distances for validation
                        bool isNewDistance = true;
                        foreach (var existing in distinctDistances)
                        {
                            if (Math.Abs(existing - dist) < DISTANCE_TOLERANCE)
                            {
                                isNewDistance = false;
                                break;
                            }
                        }
                        if (isNewDistance)
                        {
                            distinctDistances.Add(dist);
                        }
                    }
                }
            }

            if (wallThickness < 0)
                wallThickness = 0;

            distinctDistanceCount = distinctDistances.Count;
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
            // Use Identity() == 3001 (LINE_TYPE) for consistency with VBA
            if (curve == null || curve.Identity() != 3001) return false;

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

        /// <summary>
        /// Filters a face collection to only include the largest-area faces.
        /// Port of VB.NET face removal logic from CFaceCollection.vb:232-298.
        /// This removes smaller-area faces that would incorrectly contribute to cut length.
        /// </summary>
        private static List<FaceWrapper> FilterToLargestAreaFaces(List<FaceWrapper> faces)
        {
            if (faces == null || faces.Count <= 1)
                return faces ?? new List<FaceWrapper>();

            // Find the maximum area
            double maxArea = faces.Max(f => f.Area);

            // Area tolerance: faces within 1% of max area are considered "largest"
            const double AREA_TOLERANCE_RATIO = 0.01;
            double areaTolerance = maxArea * AREA_TOLERANCE_RATIO;

            // Keep only faces whose area is close to the maximum
            var result = faces.Where(f => (maxArea - f.Area) < areaTolerance).ToList();

            // If filtering removed all faces (shouldn't happen), return original
            if (result.Count == 0)
                return faces;

            ErrorHandler.DebugLog($"[TUBE] FilterToLargestAreaFaces: {faces.Count} faces -> {result.Count} " +
                $"(maxArea={maxArea * 1550003.1:F2}sq.in, tolerance={areaTolerance * 1550003.1:F4}sq.in)");

            return result;
        }

        #region Diagnostic Selection Methods (VB.NET Audit - UI/Visual Feedback)

        /// <summary>
        /// Extracts tube geometry with diagnostic information for visual debugging.
        /// Use this when a part fails classification to understand what was detected.
        /// </summary>
        public (TubeProfile Profile, TubeDiagnosticInfo Diagnostics) ExtractWithDiagnostics(IModelDoc2 model)
        {
            var diagnostics = new TubeDiagnosticInfo();
            var profile = Extract(model);

            if (model == null)
                return (profile, diagnostics);

            try
            {
                var docType = (swDocumentTypes_e)model.GetType();
                if (docType != swDocumentTypes_e.swDocPART)
                    return (profile, diagnostics);

                var partDoc = (IPartDoc)model;
                var bodiesObj = partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
                if (bodiesObj == null)
                    return (profile, diagnostics);

                var bodies = ((object[])bodiesObj).Cast<IBody2>().ToList();
                var faces = new List<FaceWrapper>();
                foreach (var body in bodies)
                {
                    var bodyFaces = body.GetFaces() as object[];
                    if (bodyFaces != null)
                        faces.AddRange(bodyFaces.Cast<IFace2>().Select(f => new FaceWrapper(f)));
                }

                if (faces.Count == 0)
                    return (profile, diagnostics);

                // Populate diagnostics based on profile type
                PopulateDiagnostics(model, faces, profile, diagnostics);
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("ExtractWithDiagnostics", ex.Message, ex);
            }

            return (profile, diagnostics);
        }

        private void PopulateDiagnostics(IModelDoc2 model, List<FaceWrapper> faces, TubeProfile profile, TubeDiagnosticInfo diagnostics)
        {
            // Find max-area faces
            double maxArea = faces.Max(f => f.Area);
            var maxAreaFaces = faces.Where(f => IsTendsToZero(f.Area - maxArea)).ToList();

            // Check for round profile
            var roundFace = maxAreaFaces.FirstOrDefault(f => f.IsRound);

            if (roundFace != null)
            {
                PopulateRoundDiagnostics(faces, roundFace, diagnostics);
            }
            else
            {
                PopulateNonRoundDiagnostics(model, faces, diagnostics);
            }
        }

        private void PopulateRoundDiagnostics(List<FaceWrapper> faces, FaceWrapper outerFace, TubeDiagnosticInfo diagnostics)
        {
            // Profile faces: outer cylinder and inner cylinder(s)
            diagnostics.ProfileFaces.Add(outerFace.Face);
            foreach (var face in faces)
            {
                if (face != outerFace && face.IsRound && outerFace.IsAxisParallelTo(face))
                    diagnostics.ProfileFaces.Add(face.Face);
            }

            // Cut length edges: outer loop edges (circles at ends)
            diagnostics.CutLengthEdges.AddRange(outerFace.GetOuterLoopEdges());

            // Hole edges
            diagnostics.HoleEdges.AddRange(outerFace.GetHoleEdges());

            // Boundary edges (same as cut length for round)
            diagnostics.BoundaryEdges.AddRange(diagnostics.CutLengthEdges);
        }

        private void PopulateNonRoundDiagnostics(IModelDoc2 model, List<FaceWrapper> faces, TubeDiagnosticInfo diagnostics)
        {
            var planarFaces = faces.Where(f => f.IsPlanar).ToList();
            if (planarFaces.Count == 0) return;

            double maxPlanarArea = planarFaces.Max(f => f.Area);
            var primaryFace = planarFaces.First(f => IsTendsToZero(f.Area - maxPlanarArea));

            // Get axis direction
            double[] startPt, endPt;
            double edgeLength;
            var axisDirection = primaryFace.GetLargestLinearEdgeDirection(out startPt, out endPt, out edgeLength);

            // Find parallel and perpendicular faces
            var primaryNormal = primaryFace.Normal;
            var facesParallelToPrimary = faces.Where(f => f.IsPlanar && f.IsNormalParallelTo(primaryFace)).ToList();
            var secondaryDir = CrossProduct(primaryNormal, axisDirection);
            var facesNormalToSecondary = faces.Where(f => f.IsPlanar && f.IsNormalParallelTo(secondaryDir)).ToList();

            // Profile faces
            var cutFaces = new List<FaceWrapper>();
            cutFaces.AddRange(facesParallelToPrimary);
            cutFaces.AddRange(facesNormalToSecondary);
            var primaryCutFaces = FilterToLargestAreaFaces(cutFaces);

            foreach (var fw in primaryCutFaces)
                diagnostics.ProfileFaces.Add(fw.Face);

            // Hole edges from all profile faces
            foreach (var face in primaryCutFaces)
                diagnostics.HoleEdges.AddRange(face.GetHoleEdges());

            // Boundary edges
            var boundaryEdges = FaceWrapper.GetBoundaryEdges(primaryCutFaces, _swApp);
            diagnostics.BoundaryEdges.AddRange(boundaryEdges);

            // Cut length edges (boundary edges perpendicular to axis + hole edges)
            foreach (var edge in boundaryEdges)
            {
                if (!IsEdgeParallelToDirection(edge, axisDirection))
                    diagnostics.CutLengthEdges.Add(edge);
            }
            diagnostics.CutLengthEdges.AddRange(diagnostics.HoleEdges);

            // If no boundary edges found, use fallback
            if (boundaryEdges.Count == 0)
            {
                foreach (var face in primaryCutFaces)
                {
                    foreach (var edge in face.GetOuterLoopEdges())
                    {
                        if (!IsEdgeParallelToDirection(edge, axisDirection))
                        {
                            if (!diagnostics.CutLengthEdges.Contains(edge))
                                diagnostics.CutLengthEdges.Add(edge);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Selects all edges that contribute to cut length calculation (Mark 1).
        /// </summary>
        public void SelectCutLengthEdges(IModelDoc2 model, TubeDiagnosticInfo diagnostics)
        {
            SelectDiagnosticElements(model, diagnostics?.CutLengthEdges, 1, "cut length edges");
        }

        /// <summary>
        /// Selects all hole edges detected on the profile (Mark 2).
        /// </summary>
        public void SelectHoleEdges(IModelDoc2 model, TubeDiagnosticInfo diagnostics)
        {
            SelectDiagnosticElements(model, diagnostics?.HoleEdges, 2, "hole edges");
        }

        /// <summary>
        /// Selects profile boundary edges (Mark 3).
        /// </summary>
        public void SelectBoundaryEdges(IModelDoc2 model, TubeDiagnosticInfo diagnostics)
        {
            SelectDiagnosticElements(model, diagnostics?.BoundaryEdges, 3, "boundary edges");
        }

        /// <summary>
        /// Selects all faces used for profile dimension calculation (Mark 4).
        /// </summary>
        public void SelectProfileFaces(IModelDoc2 model, TubeDiagnosticInfo diagnostics)
        {
            SelectDiagnosticElements(model, diagnostics?.ProfileFaces, 4, "profile faces");
        }

        /// <summary>
        /// Selects all diagnostic elements at once with different marks.
        /// Mark 1 = Cut length edges, Mark 2 = Hole edges, Mark 3 = Boundary edges, Mark 4 = Profile faces
        /// </summary>
        public void SelectAllDiagnostics(IModelDoc2 model, TubeDiagnosticInfo diagnostics)
        {
            if (model == null || diagnostics == null) return;

            model.ClearSelection2(true);
            SelectEntitiesWithMark(model, diagnostics.ProfileFaces, 4);
            SelectEntitiesWithMark(model, diagnostics.BoundaryEdges, 3);
            SelectEntitiesWithMark(model, diagnostics.HoleEdges, 2);
            SelectEntitiesWithMark(model, diagnostics.CutLengthEdges, 1);

            ErrorHandler.DebugLog($"[TUBE-DIAG] Selected all diagnostics: " +
                $"{diagnostics.ProfileFaces.Count} faces, {diagnostics.BoundaryEdges.Count} boundary, " +
                $"{diagnostics.HoleEdges.Count} holes, {diagnostics.CutLengthEdges.Count} cut length");
        }

        /// <summary>
        /// Selects a collection of COM objects in SolidWorks with a given mark, clearing the selection first.
        /// </summary>
        private void SelectDiagnosticElements<T>(IModelDoc2 model, System.Collections.Generic.IReadOnlyList<T> elements, int mark, string description)
        {
            if (model == null || elements == null) return;

            model.ClearSelection2(true);
            SelectEntitiesWithMark(model, elements, mark);
            ErrorHandler.DebugLog($"[TUBE-DIAG] Selected {elements.Count} {description}");
        }

        /// <summary>
        /// Selects entities with a specific mark value without clearing existing selection.
        /// </summary>
        private void SelectEntitiesWithMark<T>(IModelDoc2 model, System.Collections.Generic.IReadOnlyList<T> elements, int mark)
        {
            if (elements == null || elements.Count == 0) return;

            var selMgr = (ISelectionMgr)model.SelectionManager;
            var selectData = selMgr.CreateSelectData();
            selectData.Mark = mark;

            foreach (var element in elements)
            {
                var entity = (IEntity)element;
                entity.Select4(true, selectData);
            }
        }

        #endregion
    }

    /// <summary>
    /// Stores diagnostic information from tube extraction for visual debugging.
    /// Use with TubeGeometryExtractor.ExtractWithDiagnostics() and Select* methods.
    /// </summary>
    public sealed class TubeDiagnosticInfo
    {
        /// <summary>
        /// Edges that contribute to the cut length calculation.
        /// </summary>
        public List<IEdge> CutLengthEdges { get; } = new List<IEdge>();

        /// <summary>
        /// Edges that form hole perimeters on the profile.
        /// </summary>
        public List<IEdge> HoleEdges { get; } = new List<IEdge>();

        /// <summary>
        /// Boundary edges at the profile perimeter (non-round profiles only).
        /// </summary>
        public List<IEdge> BoundaryEdges { get; } = new List<IEdge>();

        /// <summary>
        /// Faces used for dimension calculation.
        /// </summary>
        public List<IFace2> ProfileFaces { get; } = new List<IFace2>();

        /// <summary>
        /// Gets a summary of the diagnostic data.
        /// </summary>
        public string GetSummary()
        {
            return $"Profile Faces: {ProfileFaces.Count}, " +
                   $"Boundary Edges: {BoundaryEdges.Count}, " +
                   $"Cut Length Edges: {CutLengthEdges.Count}, " +
                   $"Hole Edges: {HoleEdges.Count}";
        }
    }
}
