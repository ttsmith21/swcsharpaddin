using System;
using SolidWorks.Interop.sldworks;
using NM.Core.Manufacturing;

namespace NM.SwAddin.Manufacturing
{
    // Extracts perimeter/internal cut lengths and pierce counts from a flat pattern face
    public static class FlatPatternAnalyzer
    {
        private const double M_TO_IN = 39.37007874015748;

        public static CutMetrics Extract(IModelDoc2 model, IFace2 flatFace)
        {
            var cm = new CutMetrics();
            if (flatFace == null) return cm;

            try
            {
                var loopsObj = flatFace.GetLoops() as object[];
                if (loopsObj == null) return cm;
                double maxLoopLen = 0; int outerIndex = -1;
                double[] loopLens = new double[loopsObj.Length];

                for (int i = 0; i < loopsObj.Length; i++)
                {
                    var lp = loopsObj[i] as ILoop2; if (lp == null) continue;
                    var edges = lp.GetEdges() as object[]; if (edges == null) continue;
                    double sum = 0;
                    foreach (var eo in edges)
                    {
                        var e = eo as IEdge; if (e == null) continue;
                        var c = e.GetCurve() as ICurve; if (c == null) continue;
                        var p = e.GetCurveParams2() as object[];
                        double s = 0, t = 0;
                        if (p != null && p.Length >= 8)
                        {
                            double.TryParse(p[6]?.ToString(), out s);
                            double.TryParse(p[7]?.ToString(), out t);
                        }
                        double lenIn = 0;
                        try { lenIn = c.GetLength3(s, t) * M_TO_IN; } catch { }
                        sum += lenIn;
                    }
                    loopLens[i] = sum;
                    if (sum > maxLoopLen) { maxLoopLen = sum; outerIndex = i; }
                }

                cm.PierceCount = loopsObj.Length;
                cm.HoleCount = loopsObj.Length > 0 ? loopsObj.Length - 1 : 0;

                for (int i = 0; i < loopsObj.Length; i++)
                {
                    if (i == outerIndex) cm.PerimeterLengthIn += loopLens[i];
                    else cm.InternalCutLengthIn += loopLens[i];
                }
            }
            catch { }

            return cm;
        }
    }
}
