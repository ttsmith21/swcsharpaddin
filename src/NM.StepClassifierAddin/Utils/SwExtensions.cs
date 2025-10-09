using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.StepClassifierAddin.Utils
{
    internal static class SwExtensions
    {
        public static double[] ToArray3(this object p, int start = 0)
        {
            var a = p as double[]; if (a != null && a.Length >= start + 3) return new[] { a[start], a[start + 1], a[start + 2] };
            return new[] { 0.0, 0.0, 0.0 };
        }
        public static IBody2 GetFirstSolidBody(this IPartDoc part)
        {
            var bodies = part?.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
            if (bodies == null || bodies.Length == 0) return null;
            return bodies[0] as IBody2;
        }
        public static swSurfaceTypes_e GetSurfaceType(this ISurface s)
        {
            try { return (swSurfaceTypes_e)s.Identity(); } catch { return (swSurfaceTypes_e)(-1); }
        }
        public static bool TryGetPlaneNormal(this ISurface s, out double[] normal)
        {
            normal = new[] { 0.0, 0.0, 0.0 };
            try { var p = s.PlaneParams as double[]; if (p != null && p.Length >= 6) { normal = new[] { p[0], p[1], p[2] }; return true; } } catch { }
            return false;
        }
        public static bool TryGetCylinderAxis(this ISurface s, out double[] axis)
        {
            axis = new[] { 0.0, 0.0, 0.0 };
            try { var p = s.CylinderParams as double[]; if (p != null && p.Length >= 7) { axis = new[] { p[3], p[4], p[5] }; return true; } } catch { }
            return false;
        }
    }
}