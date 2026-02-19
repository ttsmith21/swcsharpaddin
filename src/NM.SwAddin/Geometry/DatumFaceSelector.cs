using System;
using System.Collections.Generic;
using System.Linq;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Geometry
{
    /// <summary>
    /// Automatically selects 3 orthogonal datum faces for DimXpert auto-dimensioning.
    /// Algorithm: largest planar face = primary, then perpendicular faces for secondary/tertiary.
    /// Falls back to linear edges when not enough perpendicular faces exist (e.g., flat blanks).
    /// </summary>
    internal sealed class DatumFaceSelector
    {
        private const double PerpendicularTolerance = 0.05; // ~3 degrees

        /// <summary>
        /// Result of automatic datum face selection.
        /// </summary>
        public sealed class DatumSelection
        {
            public bool Success { get; set; }

            /// <summary>Primary datum face (largest planar face). Mark 1 for DimXpert.</summary>
            public IFace2 PrimaryFace { get; set; }

            /// <summary>Secondary datum face (perpendicular to primary). Mark 2 for DimXpert.</summary>
            public IFace2 SecondaryFace { get; set; }

            /// <summary>Tertiary datum face (perpendicular to both primary and secondary). Mark 4.</summary>
            public IFace2 TertiaryFace { get; set; }

            /// <summary>Fallback: secondary datum edge (longest edge on primary face).</summary>
            public IEdge SecondaryEdge { get; set; }

            /// <summary>Fallback: tertiary datum edge (perpendicular to secondary edge).</summary>
            public IEdge TertiaryEdge { get; set; }

            /// <summary>Whether the edge fallback was used (flat blank with no perpendicular faces).</summary>
            public bool UsedEdgeFallback { get; set; }

            public string Message { get; set; }
        }

        /// <summary>
        /// Selects 3 orthogonal datum entities from a part's solid body geometry.
        /// </summary>
        public DatumSelection SelectDatums(IModelDoc2 model)
        {
            var result = new DatumSelection();

            try
            {
                var partDoc = model as IPartDoc;
                if (partDoc == null)
                {
                    result.Message = "Model is not a part";
                    return result;
                }

                // Get the main solid body
                var bodiesRaw = partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
                if (bodiesRaw == null)
                {
                    result.Message = "No solid bodies found";
                    return result;
                }

                var bodies = ((object[])bodiesRaw).Cast<IBody2>().ToList();
                if (bodies.Count == 0)
                {
                    result.Message = "No solid bodies found";
                    return result;
                }

                // Use the first (or largest) body
                var body = bodies[0];

                // Get all planar faces sorted by area descending
                var planarFaces = GetPlanarFacesSorted(body);
                if (planarFaces.Count == 0)
                {
                    result.Message = "No planar faces found on body";
                    return result;
                }

                // Primary datum: largest planar face
                var primary = planarFaces[0];
                result.PrimaryFace = primary.Face;
                double[] normalA = primary.Normal;

                ErrorHandler.DebugLog($"[DatumSelect] Primary face area={primary.Area * 1e6:F1}mm², normal=({normalA[0]:F3},{normalA[1]:F3},{normalA[2]:F3})");

                // Secondary datum: largest face perpendicular to primary
                for (int i = 1; i < planarFaces.Count; i++)
                {
                    double dot = Math.Abs(DotProduct(planarFaces[i].Normal, normalA));
                    if (dot < PerpendicularTolerance)
                    {
                        result.SecondaryFace = planarFaces[i].Face;
                        double[] normalB = planarFaces[i].Normal;

                        ErrorHandler.DebugLog($"[DatumSelect] Secondary face area={planarFaces[i].Area * 1e6:F1}mm², dot={dot:F4}");

                        // Tertiary datum: largest face perpendicular to both primary and secondary
                        for (int j = i + 1; j < planarFaces.Count; j++)
                        {
                            double dotA = Math.Abs(DotProduct(planarFaces[j].Normal, normalA));
                            double dotB = Math.Abs(DotProduct(planarFaces[j].Normal, normalB));
                            if (dotA < PerpendicularTolerance && dotB < PerpendicularTolerance)
                            {
                                result.TertiaryFace = planarFaces[j].Face;
                                ErrorHandler.DebugLog($"[DatumSelect] Tertiary face area={planarFaces[j].Area * 1e6:F1}mm²");
                                break;
                            }
                        }
                        break;
                    }
                }

                // Fallback: if no perpendicular faces found (flat blank), use edges
                if (result.SecondaryFace == null)
                {
                    ErrorHandler.DebugLog("[DatumSelect] No perpendicular faces found, falling back to edges");
                    SelectEdgeFallback(primary, result);
                }

                result.Success = result.PrimaryFace != null &&
                    (result.SecondaryFace != null || result.SecondaryEdge != null);
                if (result.Success)
                    result.Message = result.UsedEdgeFallback ? "Datums selected (edge fallback)" : "Datums selected";
                else
                    result.Message = "Could not find sufficient datum entities";

                return result;
            }
            catch (Exception ex)
            {
                result.Message = "Exception: " + ex.Message;
                ErrorHandler.DebugLog($"[DatumSelect] Error: {ex.Message}");
                return result;
            }
        }

        /// <summary>
        /// Gets all planar faces from a body, sorted by area descending.
        /// </summary>
        private List<FaceInfo> GetPlanarFacesSorted(IBody2 body)
        {
            var faces = new List<FaceInfo>();

            var facesRaw = body.GetFaces() as object[];
            if (facesRaw == null) return faces;

            foreach (var faceObj in facesRaw)
            {
                var face = faceObj as IFace2;
                if (face == null) continue;

                var surface = face.GetSurface() as ISurface;
                if (surface == null || !surface.IsPlane()) continue;

                var planeParams = surface.PlaneParams as double[];
                if (planeParams == null || planeParams.Length < 3) continue;

                faces.Add(new FaceInfo
                {
                    Face = face,
                    Area = face.GetArea(),
                    Normal = new double[] { planeParams[0], planeParams[1], planeParams[2] }
                });
            }

            faces.Sort((a, b) => b.Area.CompareTo(a.Area));
            return faces;
        }

        /// <summary>
        /// When no perpendicular faces exist (flat blank), select edges from the primary face
        /// as secondary and tertiary datums.
        /// </summary>
        private void SelectEdgeFallback(FaceInfo primaryFace, DatumSelection result)
        {
            result.UsedEdgeFallback = true;

            var edgesRaw = primaryFace.Face.GetEdges() as object[];
            if (edgesRaw == null || edgesRaw.Length == 0) return;

            // Find the longest linear edge
            IEdge longestEdge = null;
            double longestLength = 0;
            double[] longestDir = null;

            foreach (var edgeObj in edgesRaw)
            {
                var edge = edgeObj as IEdge;
                if (edge == null) continue;

                var curve = edge.GetCurve() as ICurve;
                if (curve == null || !curve.IsLine()) continue;

                var startV = edge.GetStartVertex() as IVertex;
                var endV = edge.GetEndVertex() as IVertex;
                if (startV == null || endV == null) continue;

                var startPt = startV.GetPoint() as double[];
                var endPt = endV.GetPoint() as double[];
                if (startPt == null || endPt == null) continue;

                double dx = endPt[0] - startPt[0];
                double dy = endPt[1] - startPt[1];
                double dz = endPt[2] - startPt[2];
                double length = Math.Sqrt(dx * dx + dy * dy + dz * dz);

                if (length > longestLength)
                {
                    longestLength = length;
                    longestEdge = edge;
                    longestDir = new double[] { dx / length, dy / length, dz / length };
                }
            }

            if (longestEdge == null) return;
            result.SecondaryEdge = longestEdge;

            // Find the longest edge perpendicular to the secondary edge
            IEdge perpEdge = null;
            double perpLength = 0;

            foreach (var edgeObj in edgesRaw)
            {
                var edge = edgeObj as IEdge;
                if (edge == null || edge == longestEdge) continue;

                var curve = edge.GetCurve() as ICurve;
                if (curve == null || !curve.IsLine()) continue;

                var startV = edge.GetStartVertex() as IVertex;
                var endV = edge.GetEndVertex() as IVertex;
                if (startV == null || endV == null) continue;

                var startPt = startV.GetPoint() as double[];
                var endPt = endV.GetPoint() as double[];
                if (startPt == null || endPt == null) continue;

                double dx = endPt[0] - startPt[0];
                double dy = endPt[1] - startPt[1];
                double dz = endPt[2] - startPt[2];
                double length = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                if (length < 1e-9) continue;

                double[] dir = { dx / length, dy / length, dz / length };
                double dot = Math.Abs(DotProduct(dir, longestDir));

                if (dot < PerpendicularTolerance && length > perpLength)
                {
                    perpLength = length;
                    perpEdge = edge;
                }
            }

            if (perpEdge != null)
                result.TertiaryEdge = perpEdge;
        }

        private static double DotProduct(double[] a, double[] b)
        {
            return a[0] * b[0] + a[1] * b[1] + a[2] * b[2];
        }

        private sealed class FaceInfo
        {
            public IFace2 Face { get; set; }
            public double Area { get; set; }
            public double[] Normal { get; set; }
        }
    }
}
