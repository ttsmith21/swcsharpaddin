using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Geometry
{
    /// <summary>
    /// Wraps a SolidWorks Face2 with geometry analysis capabilities.
    /// Ported from VB.NET ExtractData CFace class.
    /// </summary>
    internal sealed class FaceWrapper
    {
        private const double Tolerance = 1e-9;

        private readonly IFace2 _face;
        private readonly ISurface _surface;
        private readonly double _area;
        private readonly bool _isPlanar;
        private readonly bool _isRound;
        private readonly double[] _normal = new double[3];
        private readonly double[] _axis = new double[3];
        private readonly double[] _origin = new double[3];
        private readonly double _radius;

        public IFace2 Face => _face;
        public double Area => _area;
        public bool IsPlanar => _isPlanar;
        public bool IsRound => _isRound;
        public double[] Axis => _axis;
        public double[] Normal => _normal;
        public double Radius => _radius;

        public FaceWrapper(IFace2 face)
        {
            _face = face ?? throw new ArgumentNullException(nameof(face));

            _area = face.GetArea();
            _surface = (ISurface)face.GetSurface();

            if (_surface != null)
            {
                if (_surface.IsPlane())
                {
                    _isPlanar = true;
                    var normalObj = (double[])face.Normal;
                    if (normalObj != null && normalObj.Length >= 3)
                    {
                        _normal[0] = normalObj[0];
                        _normal[1] = normalObj[1];
                        _normal[2] = normalObj[2];
                    }
                }
                else if (_surface.IsCylinder())
                {
                    _isRound = true;
                    var cylParams = (double[])_surface.CylinderParams;
                    if (cylParams != null && cylParams.Length >= 7)
                    {
                        _origin[0] = cylParams[0];
                        _origin[1] = cylParams[1];
                        _origin[2] = cylParams[2];
                        _axis[0] = cylParams[3];
                        _axis[1] = cylParams[4];
                        _axis[2] = cylParams[5];
                        _radius = cylParams[6];
                    }
                }
            }
        }

        /// <summary>
        /// Gets all edges of this face.
        /// </summary>
        public List<IEdge> GetEdges()
        {
            var edgesObj = _face.GetEdges() as object[];
            if (edgesObj == null) return new List<IEdge>();
            return edgesObj.Cast<IEdge>().ToList();
        }

        /// <summary>
        /// Gets all loops on this face.
        /// </summary>
        public List<ILoop2> GetLoops()
        {
            var loopsObj = _face.GetLoops() as object[];
            if (loopsObj == null) return new List<ILoop2>();
            return loopsObj.Cast<ILoop2>().ToList();
        }

        /// <summary>
        /// Gets inner loops (holes) on this face.
        /// </summary>
        public List<ILoop2> GetHoles()
        {
            return GetLoops().Where(loop => !loop.IsOuter()).ToList();
        }

        /// <summary>
        /// Gets outer loops on this face.
        /// </summary>
        public List<ILoop2> GetOuterLoops()
        {
            return GetLoops().Where(loop => loop.IsOuter()).ToList();
        }

        /// <summary>
        /// Gets all edges from outer loops.
        /// </summary>
        public List<IEdge> GetOuterLoopEdges()
        {
            var edges = new List<IEdge>();
            foreach (var loop in GetOuterLoops())
            {
                var loopEdges = loop.GetEdges() as object[];
                if (loopEdges != null)
                {
                    edges.AddRange(loopEdges.Cast<IEdge>());
                }
            }
            return edges;
        }

        /// <summary>
        /// Gets all edges from hole loops.
        /// </summary>
        public List<IEdge> GetHoleEdges()
        {
            var edges = new List<IEdge>();
            foreach (var loop in GetHoles())
            {
                var loopEdges = loop.GetEdges() as object[];
                if (loopEdges != null)
                {
                    edges.AddRange(loopEdges.Cast<IEdge>());
                }
            }
            return edges;
        }

        /// <summary>
        /// Checks if this face is the same as another face.
        /// </summary>
        public bool IsSame(FaceWrapper other, ISldWorks swApp)
        {
            if (other == null) return false;
            return IsSame(other.Face, swApp);
        }

        /// <summary>
        /// Checks if this face is the same as another Face2.
        /// </summary>
        public bool IsSame(IFace2 otherFace, ISldWorks swApp)
        {
            if (otherFace == null || swApp == null) return false;
            return swApp.IsSame(_face, otherFace) == (int)swObjectEquality.swObjectSame;
        }

        /// <summary>
        /// Checks if another face's axis is parallel to this face's axis.
        /// </summary>
        public bool IsAxisParallelTo(FaceWrapper other)
        {
            if (other == null || !IsRound || !other.IsRound) return false;
            return IsVectorParallel(_axis, other.Axis);
        }

        /// <summary>
        /// Checks if another face's normal is parallel to this face's normal.
        /// </summary>
        public bool IsNormalParallelTo(FaceWrapper other)
        {
            if (other == null || !IsPlanar || !other.IsPlanar) return false;
            return IsVectorParallel(_normal, other.Normal);
        }

        /// <summary>
        /// Checks if a direction is parallel to this face's normal.
        /// </summary>
        public bool IsNormalParallelTo(double[] direction)
        {
            if (direction == null || !IsPlanar) return false;
            return IsVectorParallel(_normal, direction);
        }

        /// <summary>
        /// Selects this face in the model.
        /// </summary>
        public bool SelectFace(SelectData selectData, bool append)
        {
            var entity = (IEntity)_face;
            return entity.Select4(append, selectData);
        }

        /// <summary>
        /// Gets the direction and length of the largest linear edge on this face.
        /// </summary>
        public double[] GetLargestLinearEdgeDirection(out double[] startPoint, out double[] endPoint, out double length)
        {
            startPoint = new double[3];
            endPoint = new double[3];
            length = 0;
            var direction = new double[3];

            var edges = GetEdges();
            double maxLength = 0;

            foreach (var edge in edges)
            {
                var curve = (ICurve)edge.GetCurve();
                // Use Identity() == 3001 (LINE_TYPE) for consistency with VBA
                if (curve == null || curve.Identity() != 3001) continue;

                double start = 0, end = 0;
                bool isClosed = false, isPeriodic = false;
                if (!curve.GetEndParams(out start, out end, out isClosed, out isPeriodic)) continue;

                double edgeLength = curve.GetLength3(start, end);
                if (edgeLength > maxLength)
                {
                    maxLength = edgeLength;
                    length = edgeLength;

                    var lineParams = (double[])curve.LineParams;
                    if (lineParams != null && lineParams.Length >= 6)
                    {
                        direction[0] = lineParams[3];
                        direction[1] = lineParams[4];
                        direction[2] = lineParams[5];
                    }

                    int numPoints = 0;
                    var startEval = (double[])curve.Evaluate2(start, numPoints);
                    var endEval = (double[])curve.Evaluate2(end, numPoints);
                    if (startEval != null && startEval.Length >= 3)
                    {
                        startPoint[0] = startEval[0];
                        startPoint[1] = startEval[1];
                        startPoint[2] = startEval[2];
                    }
                    if (endEval != null && endEval.Length >= 3)
                    {
                        endPoint[0] = endEval[0];
                        endPoint[1] = endEval[1];
                        endPoint[2] = endEval[2];
                    }
                }
            }

            return direction;
        }

        /// <summary>
        /// Calculates material length from tessellation triangles along the axis.
        /// </summary>
        public double GetMaterialLength(out double[] startPoint, out double[] endPoint, double[] axisDirection = null)
        {
            startPoint = new double[3];
            endPoint = new double[3];

            var tessTriangles = _face.GetTessTriangles(true) as double[];
            if (tessTriangles == null || tessTriangles.Length < 9) return 0;

            double xMin = double.MaxValue, xMax = double.MinValue;
            double yMin = double.MaxValue, yMax = double.MinValue;
            double zMin = double.MaxValue, zMax = double.MinValue;

            // Process tessellation triangles (9 values per triangle: 3 vertices x 3 coords)
            int numTriangles = tessTriangles.Length / 9;
            for (int i = 0; i < numTriangles; i++)
            {
                int baseIdx = i * 9;
                for (int v = 0; v < 3; v++)
                {
                    double x = tessTriangles[baseIdx + v * 3 + 0];
                    double y = tessTriangles[baseIdx + v * 3 + 1];
                    double z = tessTriangles[baseIdx + v * 3 + 2];

                    xMin = Math.Min(xMin, x); xMax = Math.Max(xMax, x);
                    yMin = Math.Min(yMin, y); yMax = Math.Max(yMax, y);
                    zMin = Math.Min(zMin, z); zMax = Math.Max(zMax, z);
                }
            }

            startPoint[0] = xMin; startPoint[1] = yMin; startPoint[2] = zMin;
            endPoint[0] = xMax; endPoint[1] = yMax; endPoint[2] = zMax;

            double[] axis = axisDirection ?? _axis;
            double materialLength;

            // Calculate length along axis direction
            if (IsTendsToZero(Math.Abs(axis[0]) - 1))
            {
                materialLength = xMax - xMin;
            }
            else if (IsTendsToZero(Math.Abs(axis[1]) - 1))
            {
                materialLength = yMax - yMin;
            }
            else if (IsTendsToZero(Math.Abs(axis[2]) - 1))
            {
                materialLength = zMax - zMin;
            }
            else
            {
                // General case: project onto axis
                double axisMag = Math.Sqrt(axis[0] * axis[0] + axis[1] * axis[1] + axis[2] * axis[2]);
                if (axisMag < Tolerance) return 0;

                materialLength = (Math.Abs(axis[0]) * (xMax - xMin) +
                                  Math.Abs(axis[1]) * (yMax - yMin) +
                                  Math.Abs(axis[2]) * (zMax - zMin)) / axisMag;
            }

            return materialLength;
        }

        /// <summary>
        /// Gets the adjacent face of an edge that is not this face.
        /// </summary>
        public IFace2 GetAdjacentFace(IEdge edge, ISldWorks swApp)
        {
            if (edge == null) return null;

            var faces = edge.GetTwoAdjacentFaces2() as object[];
            if (faces == null || faces.Length < 2) return null;

            var face0 = faces[0] as IFace2;
            var face1 = faces[1] as IFace2;

            if (face0 != null && swApp.IsSame(_face, face0) == (int)swObjectEquality.swObjectSame)
                return face1;
            return face0;
        }

        private static bool IsVectorParallel(double[] v1, double[] v2)
        {
            if (v1 == null || v2 == null || v1.Length < 3 || v2.Length < 3) return false;

            // Vectors are parallel if they point in the same or opposite direction
            bool sameDir = IsTendsToZero(v1[0] - v2[0]) && IsTendsToZero(v1[1] - v2[1]) && IsTendsToZero(v1[2] - v2[2]);
            bool oppDir = IsTendsToZero(v1[0] + v2[0]) && IsTendsToZero(v1[1] + v2[1]) && IsTendsToZero(v1[2] + v2[2]);
            return sameDir || oppDir;
        }

        private static bool IsTendsToZero(double val)
        {
            return Math.Abs(val) < Tolerance;
        }

        #region Edge Filtering (VB.NET Audit Resolution)

        /// <summary>
        /// Gets both faces adjacent to an edge.
        /// Port of VB.NET GetAdjacentFaceOfThisEdgeOtherThanMe pattern.
        /// </summary>
        public static (IFace2 face1, IFace2 face2) GetAdjacentFaces(IEdge edge)
        {
            if (edge == null) return (null, null);

            var faces = edge.GetTwoAdjacentFaces2() as object[];
            if (faces == null || faces.Length < 2)
                return (null, null);

            return ((IFace2)faces[0], (IFace2)faces[1]);
        }

        /// <summary>
        /// Gets boundary edges - edges where one adjacent face is in the collection and one is not.
        /// Both adjacent faces must be planar.
        /// Port of VB.NET GetAllEdgesWhoseAdjacentFacesArePlanarAndWhoseOtherAdjacentFaceIsNotInThisCollection.
        /// </summary>
        /// <param name="faceCollection">Collection of faces defining the profile</param>
        /// <param name="swApp">SolidWorks application for face comparison</param>
        /// <returns>List of boundary edges</returns>
        public static List<IEdge> GetBoundaryEdges(List<FaceWrapper> faceCollection, ISldWorks swApp)
        {
            if (faceCollection == null || faceCollection.Count == 0 || swApp == null)
                return new List<IEdge>();

            var result = new List<IEdge>();
            var processedEdges = new HashSet<object>(); // Track processed edges by reference

            foreach (var fw in faceCollection)
            {
                foreach (var edge in fw.GetOuterLoopEdges())
                {
                    // Skip if already processed
                    if (processedEdges.Contains(edge)) continue;
                    processedEdges.Add(edge);

                    var (face1, face2) = GetAdjacentFaces(edge);
                    if (face1 == null || face2 == null) continue;

                    // Both faces must be planar
                    var surf1 = (ISurface)face1.GetSurface();
                    var surf2 = (ISurface)face2.GetSurface();
                    if (surf1 == null || surf2 == null) continue;
                    if (!surf1.IsPlane() || !surf2.IsPlane()) continue;

                    // Check if one face is in collection, one is not
                    bool f1InCollection = IsFaceInCollection(face1, faceCollection, swApp);
                    bool f2InCollection = IsFaceInCollection(face2, faceCollection, swApp);

                    // XOR - exactly one must be in collection
                    if (f1InCollection != f2InCollection)
                    {
                        result.Add(edge);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Checks if a face is in the face wrapper collection.
        /// </summary>
        private static bool IsFaceInCollection(IFace2 face, List<FaceWrapper> collection, ISldWorks swApp)
        {
            if (face == null || collection == null) return false;

            foreach (var fw in collection)
            {
                if (swApp.IsSame(face, fw.Face) == (int)swObjectEquality.swObjectSame)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Checks if edge endpoints lie on any face in the collection.
        /// Port of VB.NET DoesStartAndEndPointsOfThisEdgeLieOnFacesFromThisFaceCollection.
        /// Uses vertex-to-face proximity check.
        /// </summary>
        /// <param name="edge">Edge to check</param>
        /// <param name="faces">Collection of faces</param>
        /// <param name="tolerance">Distance tolerance for point-on-face check (default 1e-6 meters)</param>
        /// <returns>True if both endpoints lie on faces in the collection</returns>
        public static bool DoEdgeEndpointsLieOnFaces(IEdge edge, List<FaceWrapper> faces, double tolerance = 1e-6)
        {
            if (edge == null || faces == null || faces.Count == 0)
                return false;

            var startVertex = (IVertex)edge.GetStartVertex();
            var endVertex = (IVertex)edge.GetEndVertex();

            if (startVertex == null || endVertex == null) return false;

            var startPoint = (double[])startVertex.GetPoint();
            var endPoint = (double[])endVertex.GetPoint();

            if (startPoint == null || endPoint == null) return false;

            bool startOnFace = IsPointOnAnyFace(startPoint, faces, tolerance);
            bool endOnFace = IsPointOnAnyFace(endPoint, faces, tolerance);

            return startOnFace && endOnFace;
        }

        /// <summary>
        /// Checks if a point lies on any face in the collection using closest point distance.
        /// </summary>
        private static bool IsPointOnAnyFace(double[] point, List<FaceWrapper> faces, double tolerance)
        {
            if (point == null || point.Length < 3) return false;

            foreach (var fw in faces)
            {
                // Use face's GetClosestPointOn to check if point is on face
                var closestPoint = fw.Face.GetClosestPointOn(point[0], point[1], point[2]) as double[];
                if (closestPoint != null && closestPoint.Length >= 3)
                {
                    double dx = closestPoint[0] - point[0];
                    double dy = closestPoint[1] - point[1];
                    double dz = closestPoint[2] - point[2];
                    double distSq = dx * dx + dy * dy + dz * dz;

                    if (distSq < tolerance * tolerance)
                        return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Gets all edges that share a loop with the given edge on the specified face.
        /// Port of VB.NET GetEdgesWhichAreInLoopWithThisEdge.
        /// </summary>
        /// <param name="edge">The reference edge</param>
        /// <param name="face">The face containing the edge</param>
        /// <returns>List of other edges in the same loop</returns>
        public static List<IEdge> GetEdgesInSameLoop(IEdge edge, IFace2 face)
        {
            var result = new List<IEdge>();
            if (edge == null || face == null) return result;

            var loops = face.GetLoops() as object[];
            if (loops == null) return result;

            foreach (ILoop2 loop in loops.Cast<ILoop2>())
            {
                var loopEdges = loop.GetEdges() as object[];
                if (loopEdges == null) continue;

                var edgeList = loopEdges.Cast<IEdge>().ToList();

                // Check if this loop contains the target edge
                bool containsEdge = edgeList.Any(e => ReferenceEquals(e, edge));

                if (containsEdge)
                {
                    // Return all other edges in this loop
                    result.AddRange(edgeList.Where(e => !ReferenceEquals(e, edge)));
                    break; // Edge can only be in one loop per face
                }
            }

            return result;
        }

        #endregion
    }
}
