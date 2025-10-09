using System;
using SolidWorks.Interop.sldworks;
using NM.StepClassifierAddin.Utils;

namespace NM.StepClassifierAddin.Classification
{
    internal static class Slicer
    {
        public static bool VerifyConstantSection(ISldWorks app, IModelDoc2 model, IBody2 body, double[] axisUnit, double sectionAreaTol)
        {
            if (app == null || model == null || body == null || axisUnit == null) return false;
            double[] axis = Math3D.Normalize(axisUnit);

            double minx, miny, minz, maxx, maxy, maxz;
            body.GetExtremePoint(-axis[0], -axis[1], -axis[2], out minx, out miny, out minz);
            body.GetExtremePoint(axis[0], axis[1], axis[2], out maxx, out maxy, out maxz);
            double L = Math.Sqrt((maxx - minx) * (maxx - minx) + (maxy - miny) * (maxy - miny) + (maxz - minz) * (maxz - minz));
            if (L <= 0) return false;

            var pA = new[] { minx + axis[0] * 0.35 * L, miny + axis[1] * 0.35 * L, minz + axis[2] * 0.35 * L };
            var pB = new[] { minx + axis[0] * 0.65 * L, miny + axis[1] * 0.65 * L, minz + axis[2] * 0.65 * L };

            var mod = app.IGetModeler(); if (mod == null) return false;

            double areaA = SliceAndMeasure(mod, body, pA, axis);
            double areaB = SliceAndMeasure(mod, body, pB, axis);
            if (areaA <= 0 || areaB <= 0) return false;
            double rel = Math.Abs(areaA - areaB) / Math.Max(areaA, areaB);
            return rel <= sectionAreaTol;
        }

        private static double SliceAndMeasure(IModeler mod, IBody2 body, double[] point, double[] normal)
        {
            IBody2 sheet = null; object facesObj = null; Body2 outBodies = null; try
            {
                // Use CreatePlanarSurface2 with origin, normal, refdir
                object origin = new object[] { point[0], point[1], point[2] };
                object nrm = new object[] { normal[0], normal[1], normal[2] };
                object refdir = new object[] { 1.0, 0.0, 0.0 };
                var surf = mod.CreatePlanarSurface2(origin, nrm, refdir);
                object uvRange = null;
                sheet = mod.CreateSheetFromSurface(surf, uvRange) as IBody2;
                if (sheet == null) return 0.0;

                facesObj = body.ISectionBySheet((Body2)sheet, 1, ref outBodies);
                var faces = facesObj as object[];
                if (faces == null || faces.Length == 0) return 0.0;

                double totalArea = 0.0;
                foreach (var o in faces)
                {
                    var face = o as IFace2; if (face == null) continue;
                    double area = 0.0; try { area = face.GetArea(); } catch { }
                    totalArea += Math.Abs(area);
                }
                return totalArea;
            }
            catch { return 0.0; }
            finally
            {
                try { if (sheet != null) ((IDisposable)sheet).Dispose(); } catch { }
            }
        }
    }
}