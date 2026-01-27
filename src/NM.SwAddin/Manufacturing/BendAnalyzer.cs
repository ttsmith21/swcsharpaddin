using System;
using NM.Core.Manufacturing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Manufacturing
{
    public static class BendAnalyzer
    {
        private const double M_TO_IN = 39.37007874015748;

        public static BendInfo AnalyzeBends(IModelDoc2 model)
        {
            var info = new BendInfo();
            if (model == null) return info;

            try
            {
                IFeature feat = model.FirstFeature() as IFeature;
                bool hasUp = false, hasDown = false;
                while (feat != null)
                {
                    string type = feat.GetTypeName2() ?? string.Empty;
                    if (type == "OneBend" || type == "SketchBend" || type == "EdgeFlange")
                    {
                        // Skip suppressed
                        bool suppressed = false;
                        try { suppressed = feat.IsSuppressed(); } catch { }
                        if (suppressed) { feat = feat.GetNextFeature() as IFeature; continue; }

                        info.Count++;
                        try
                        {
                            object defObj = feat.GetDefinition();
                            if (defObj is IOneBendFeatureData one)
                            {
                                double rIn = (one.BendRadius) * M_TO_IN;
                                if (rIn > info.MaxRadiusIn) info.MaxRadiusIn = rIn;
                                double ang = 0.0;
                                try { ang = one.BendAngle; } catch { }
                                if (ang > 0) hasUp = true; else if (ang < 0) hasDown = true;

                                // Heuristic for longest bend if we can't access sketch/edge: use radius as proxy
                                double approxLen = rIn * 3.14159; // half circumference proxy
                                if (approxLen > info.LongestBendIn) info.LongestBendIn = approxLen;
                            }
                        }
                        catch { }
                    }
                    feat = feat.GetNextFeature() as IFeature;
                }
                info.NeedsFlip = hasUp && hasDown;
            }
            catch { }

            return info;
        }
    }
}
