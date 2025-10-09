using System;
using System.Reflection;
using SolidWorks.Interop.sldworks;
using NM.StepClassifierAddin.Utils;

namespace NM.StepClassifierAddin.Classification
{
    internal static class GeometryScanner
    {
        /// <summary>
        /// Scan body once to compute principal axes, area tallies, and oriented spans.
        /// </summary>
        public static PrepassMetrics Scan(IBody2 body, IMassProperty mp)
        {
            if (body == null) throw new ArgumentNullException(nameof(body));
            if (mp == null) throw new ArgumentNullException(nameof(mp));

            var m = new PrepassMetrics { Body = body };

            // Principal axes via interop: try property/indexer or fallback method on either IMassProperty or IMassProperty2
            try
            {
                object mp2 = mp; // we will reflect on whatever COM object we have
                double[] axes = null;
                var prop = mp2.GetType().GetProperty("PrincipalAxesOfInertia");
                if (prop != null)
                {
                    var idx = prop.GetIndexParameters();
                    if (idx != null && idx.Length == 1)
                    {
                        var obj = prop.GetValue(mp2, new object[] { 0 });
                        axes = obj as double[];
                    }
                }
                if (axes == null)
                {
                    var mi = mp2.GetType().GetMethod("PrincipalAxesOfInertia");
                    if (mi != null)
                    {
                        var pars = mi.GetParameters();
                        if (pars.Length >= 3)
                        {
                            object v1 = null, v2 = null, v3 = null, p1 = null, p2 = null, p3 = null;
                            object[] args = new object[] { v1, v2, v3, p1, p2, p3 };
                            mi.Invoke(mp2, args);
                            var a1 = (args[0] as double[]) ?? new double[] { 1, 0, 0 };
                            var a2 = (args[1] as double[]) ?? new double[] { 0, 1, 0 };
                            var a3 = (args[2] as double[]) ?? new double[] { 0, 0, 1 };
                            m.Axes[0].Axis = Math3D.Normalize(new[] { a1[0], a1[1], a1[2] });
                            m.Axes[1].Axis = Math3D.Normalize(new[] { a2[0], a2[1], a2[2] });
                            m.Axes[2].Axis = Math3D.Normalize(new[] { a3[0], a3[1], a3[2] });
                        }
                    }
                }
                if (m.Axes[0].Axis[0] == 0 && m.Axes[0].Axis[1] == 0 && m.Axes[0].Axis[2] == 0)
                {
                    if (axes != null && axes.Length >= 9)
                    {
                        m.Axes[0].Axis = Math3D.Normalize(new[] { axes[0], axes[1], axes[2] });
                        m.Axes[1].Axis = Math3D.Normalize(new[] { axes[3], axes[4], axes[5] });
                        m.Axes[2].Axis = Math3D.Normalize(new[] { axes[6], axes[7], axes[8] });
                    }
                    else
                    {
                        m.Axes[0].Axis = new[] { 1.0, 0.0, 0.0 };
                        m.Axes[1].Axis = new[] { 0.0, 1.0, 0.0 };
                        m.Axes[2].Axis = new[] { 0.0, 0.0, 1.0 };
                    }
                }
            }
            catch { m.Axes[0].Axis = new[] { 1.0, 0.0, 0.0 }; m.Axes[1].Axis = new[] { 0.0, 1.0, 0.0 }; m.Axes[2].Axis = new[] { 0.0, 0.0, 1.0 }; }

            var faces = body.GetFaces() as object[];
            if (faces != null)
            {
                foreach (var fo in faces)
                {
                    var f = fo as IFace2; if (f == null) continue;
                    double area = 0.0; try { area = f.GetArea(); } catch { }
                    m.ATotal += area;
                    var surf = f.IGetSurface(); if (surf == null) continue;

                    bool isPlane = false, isCyl = false, isCone = false;
                    try { isPlane = surf.IsPlane(); } catch { }
                    try { isCyl = surf.IsCylinder(); } catch { }
                    try { isCone = surf.IsCone(); } catch { }
                    if (isPlane || isCyl || isCone) m.ADevelopable += area;

                    for (int i = 0; i < 3; i++)
                    {
                        var v = m.Axes[i].Axis;
                        if (isPlane)
                        {
                            if (!surf.TryGetPlaneNormal(out var n)) continue;
                            n = Math3D.Normalize(n);
                            double ad = Math.Abs(Math3D.Dot(n, v));
                            if (ad <= Thresholds.NORMAL_PERP_TOL) m.Axes[i].SideArea += area;
                            if (ad >= 1.0 - Thresholds.CAP_ALIGN_TOL) m.Axes[i].CapArea += area;
                        }
                        else if (isCyl)
                        {
                            if (surf.TryGetCylinderAxis(out var a))
                            {
                                a = Math3D.Normalize(a);
                                double ad = Math.Abs(Math3D.Dot(a, v));
                                if (ad >= 1.0 - Thresholds.ALIGN_TOL) m.Axes[i].SideArea += area;
                            }
                        }
                    }
                }
            }

            // Oriented spans using extreme points (use out params)
            try
            {
                for (int i = 0; i < 3; i++)
                {
                    var dir = m.Axes[i].Axis;
                    double minx, miny, minz, maxx, maxy, maxz;
                    body.GetExtremePoint(-dir[0], -dir[1], -dir[2], out minx, out miny, out minz);
                    body.GetExtremePoint(dir[0], dir[1], dir[2], out maxx, out maxy, out maxz);
                    double dx = maxx - minx, dy = maxy - miny, dz = maxz - minz;
                    double span = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                    if (i == 0) m.Spans.L1 = span; else if (i == 1) m.Spans.L2 = span; else m.Spans.L3 = span;
                }
                // Sort descending
                double[] arr = new[] { m.Spans.L1, m.Spans.L2, m.Spans.L3 };
                Array.Sort(arr); Array.Reverse(arr);
                m.Spans.L1 = arr[0]; m.Spans.L2 = arr[1]; m.Spans.L3 = arr[2];
            }
            catch { }

            return m;
        }
    }
}