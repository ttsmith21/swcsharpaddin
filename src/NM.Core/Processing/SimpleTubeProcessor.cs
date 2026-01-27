using System;
using System.Linq;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.Core.Processing
{
    /// <summary>
    /// Tube processing implementation - initial geometry extraction implemented.
    /// All supporting services exist; this extracts basic tube geometry (OD, wall, length, axis).
    /// </summary>
    public class SimpleTubeProcessor
    {
        private const string LogPrefix = "[TUBE]";
        private readonly ISldWorks _swApp;

        public SimpleTubeProcessor() { }
        public SimpleTubeProcessor(ISldWorks swApp) { _swApp = swApp; }

        public string Name => "Tube";

        /// <summary>
        /// Returns true if the model appears to be a simple tube (hollow cylinder with planar end rings).
        /// </summary>
        public bool CanProcess(IModelDoc2 model)
        {
            try
            {
                var g = ExtractTubeGeometry(model);
                bool ok = g != null && g.OuterDiameter > 0 && g.WallThickness > 0 && g.Length > 0;
                ErrorHandler.DebugLog($"{LogPrefix} CanProcess ? {(ok ? "YES" : "NO")} (OD={g?.OuterDiameter:F3}in, Wall={g?.WallThickness:F3}in, L={g?.Length:F3}in)");
                return ok;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("SimpleTubeProcessor.CanProcess", "Detection failed", ex, ErrorHandler.LogLevel.Warning);
                return false;
            }
        }

        /// <summary>
        /// Public helper for callers that need geometry without running full processing.
        /// Returns null if geometry is not detected.
        /// </summary>
        public TubeGeometry TryGetGeometry(IModelDoc2 model)
        {
            return ExtractTubeGeometry(model);
        }

        /// <summary>
        /// Uses geometry extraction; downstream services (schedule/material/cutting) to be integrated next.
        /// </summary>
        public bool Process(IModelDoc2 model, ProcessingOptions options)
        {
            return Process(null, model, options);
        }

        /// <summary>
        /// Process with ModelInfo context for custom property updates.
        /// </summary>
        public bool Process(ModelInfo info, IModelDoc2 model, ProcessingOptions options)
        {
            ErrorHandler.PushCallStack("SimpleTubeProcessor.Process");
            try
            {
                var geom = ExtractTubeGeometry(model);
                if (geom == null)
                {
                    ErrorHandler.HandleError("SimpleTubeProcessor.Process", "Not a tube (geometry not detected)", null, ErrorHandler.LogLevel.Warning);
                    return false;
                }

                ErrorHandler.DebugLog($"{LogPrefix} OD={geom.OuterDiameter:F3}in, Wall={geom.WallThickness:F3}in, L={geom.Length:F3}in, Axis=({geom.Axis[0]:F3},{geom.Axis[1]:F3},{geom.Axis[2]:F3})");

                // Write properties to ModelInfo if provided
                if (info?.CustomProperties != null)
                {
                    info.CustomProperties.SetPropertyValue("TubeOD", geom.OuterDiameter.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture), CustomPropertyType.Number);
                    info.CustomProperties.SetPropertyValue("TubeWall", geom.WallThickness.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture), CustomPropertyType.Number);
                    info.CustomProperties.SetPropertyValue("TubeLength", geom.Length.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture), CustomPropertyType.Number);
                    info.CustomProperties.IsTube = true;
                }

                // TODO(vNext): Use PipeScheduleService, TubeMaterialCodeGenerator, TubeCuttingParameterService, ExternalStart
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("SimpleTubeProcessor.Process", "Tube processing failed", ex, ErrorHandler.LogLevel.Error);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        #region Geometry Extraction
        private TubeGeometry ExtractTubeGeometry(IModelDoc2 model)
        {
            const string proc = "SimpleTubeProcessor.ExtractTubeGeometry";
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (model == null) return null;
                var part = model as IPartDoc;
                if (part == null) return null;

                var bodies = part.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
                if (bodies == null || bodies.Length == 0) return null;
                var body = bodies[0] as IBody2;
                if (body == null) return null;

                var faces = body.GetFaces() as object[];
                if (faces == null || faces.Length == 0) return null;

                // Collect cylinder radii and a representative axis
                double maxR_m = 0.0, minR_m = double.MaxValue;
                double[] axis = null; // normalized
                double[] axisOrigin = null;
                int cylCount = 0;

                foreach (var f in faces)
                {
                    var face = f as IFace2; if (face == null) continue;
                    var surf = face.IGetSurface(); if (surf == null) continue;
                    if (surf.IsCylinder())
                    {
                        var p = surf.CylinderParams as double[];
                        if (p == null || p.Length < 7) continue;
                        var a = new[] { p[3], p[4], p[5] };
                        var len = Math.Sqrt(a[0] * a[0] + a[1] * a[1] + a[2] * a[2]);
                        if (len < 1e-9) continue;
                        var n = new[] { a[0] / len, a[1] / len, a[2] / len };
                        var r = Math.Abs(p[6]);

                        // Initialize axis with first cylinder
                        if (axis == null)
                        {
                            axis = n; axisOrigin = new[] { p[0], p[1], p[2] };
                        }
                        else
                        {
                            // Ensure same orientation (flip if necessary)
                            var dp = axis[0] * n[0] + axis[1] * n[1] + axis[2] * n[2];
                            if (dp < 0) { n[0] = -n[0]; n[1] = -n[1]; n[2] = -n[2]; }
                        }

                        if (r > maxR_m) maxR_m = r;
                        if (r < minR_m) minR_m = r;
                        cylCount++;
                    }
                }

                if (cylCount == 0 || axis == null) return null;

                // Estimate length: project all vertices onto axis and compute span
                double minT = double.MaxValue, maxT = double.MinValue;
                var vertsObj = body.GetVertices() as object[];
                if (vertsObj == null || vertsObj.Length == 0)
                {
                    // Fallback: use face vertices
                    foreach (var f in faces)
                    {
                        var face = f as IFace2; if (face == null) continue;
                        var loops = face.GetLoops() as object[]; if (loops == null) continue;
                        foreach (var l in loops)
                        {
                            var loop = l as ILoop2; if (loop == null) continue;
                            var edges = loop.GetEdges() as object[]; if (edges == null) continue;
                            foreach (var e in edges)
                            {
                                var edge = e as IEdge; if (edge == null) continue;
                                // Get start and end vertices
                                var sv = edge.IGetStartVertex();
                                var ev = edge.IGetEndVertex();
                                foreach (var vv in new[] { sv, ev })
                                {
                                    if (vv == null) continue;
                                    var pt = vv.GetPoint() as double[]; if (pt == null || pt.Length < 3) continue;
                                    var t = Dot(Sub(pt, axisOrigin), axis);
                                    if (t < minT) minT = t;
                                    if (t > maxT) maxT = t;
                                }
                            }
                        }
                    }
                }
                else
                {
                    foreach (var v in vertsObj)
                    {
                        var vv = v as IVertex; if (vv == null) continue;
                        var pt = vv.GetPoint() as double[]; if (pt == null || pt.Length < 3) continue;
                        var t = Dot(Sub(pt, axisOrigin), axis);
                        if (t < minT) minT = t;
                        if (t > maxT) maxT = t;
                    }
                }

                double length_m = (maxT > minT && maxT < double.MaxValue && minT > double.MinValue) ? (maxT - minT) : 0.0;

                // Require two cylinders for hollow tube; otherwise it's a solid rod
                if (cylCount < 2) return null;

                var od_m = 2 * maxR_m;
                var id_m = 2 * minR_m;
                var wall_m = (maxR_m - minR_m);

                if (od_m <= 0 || wall_m <= 1e-6 || length_m <= 0) return null;

                // Convert to inches for TubeGeometry DTO
                const double M_TO_IN = 39.37007874015748;
                var geom = new TubeGeometry
                {
                    OuterDiameter = od_m * M_TO_IN,
                    WallThickness = wall_m * M_TO_IN,
                    Length = length_m * M_TO_IN,
                    Axis = axis
                };
                return geom;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Geometry extraction failed", ex, ErrorHandler.LogLevel.Warning);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        private static double[] Sub(double[] a, double[] b) => new[] { a[0] - b[0], a[1] - b[1], a[2] - b[2] };
        private static double Dot(double[] a, double[] b) => a[0] * b[0] + a[1] * b[1] + a[2] * b[2];
        #endregion
    }

    /// <summary>
    /// Data structure for tube geometry - ready for use once extraction works.
    /// Units: inches for OD/Wall/Length.
    /// </summary>
    public class TubeGeometry
    {
        public double OuterDiameter { get; set; }  // inches
        public double WallThickness { get; set; }   // inches
        public double Length { get; set; }          // inches
        public double[] Axis { get; set; }          // normalized direction vector [x, y, z]
    }
}
