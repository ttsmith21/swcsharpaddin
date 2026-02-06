using System;
using NM.Core.Manufacturing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Manufacturing
{
    public static class BendAnalyzer
    {
        private const double M_TO_IN = 39.37007874015748;

        // Feature types that contain bend sub-features (from InsertBends2)
        private static readonly string[] BendContainerTypes = { "ProcessBends", "FlattenBends" };

        /// <summary>
        /// Analyze bend features in the model feature tree.
        /// Searches both top-level features and sub-features under ProcessBends/FlattenBends.
        /// </summary>
        /// <param name="model">The SolidWorks model document.</param>
        /// <param name="countSuppressed">If true, count suppressed bends (e.g. when model is flattened). Default false.</param>
        public static BendInfo AnalyzeBends(IModelDoc2 model, bool countSuppressed = false)
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

                    // Check if this feature itself is a bend
                    if (IsBendFeature(type))
                    {
                        CountBend(feat, info, countSuppressed, ref hasUp, ref hasDown);
                    }

                    // Check sub-features under bend container features (ProcessBends, FlattenBends)
                    // OneBend/SketchBend features from InsertBends2 are sub-features, not top-level
                    if (IsBendContainer(type))
                    {
                        var subFeat = feat.GetFirstSubFeature() as IFeature;
                        while (subFeat != null)
                        {
                            string subType = subFeat.GetTypeName2() ?? string.Empty;
                            if (IsBendFeature(subType))
                            {
                                CountBend(subFeat, info, countSuppressed, ref hasUp, ref hasDown);
                            }
                            subFeat = subFeat.GetNextSubFeature() as IFeature;
                        }
                    }

                    feat = feat.GetNextFeature() as IFeature;
                }
                info.NeedsFlip = hasUp && hasDown;
            }
            catch { }

            return info;
        }

        private static bool IsBendFeature(string typeName)
        {
            return typeName == "OneBend" || typeName == "SketchBend" || typeName == "EdgeFlange";
        }

        private static bool IsBendContainer(string typeName)
        {
            for (int i = 0; i < BendContainerTypes.Length; i++)
            {
                if (typeName == BendContainerTypes[i]) return true;
            }
            return false;
        }

        private static void CountBend(IFeature feat, BendInfo info, bool countSuppressed, ref bool hasUp, ref bool hasDown)
        {
            bool suppressed = false;
            try { suppressed = feat.IsSuppressed(); } catch { }
            if (suppressed && !countSuppressed) return;

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
    }
}
