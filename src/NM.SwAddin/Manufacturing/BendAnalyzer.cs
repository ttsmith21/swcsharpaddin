using System;
using NM.Core;
using NM.Core.Manufacturing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using static NM.Core.Constants.UnitConversions;

namespace NM.SwAddin.Manufacturing
{
    public static class BendAnalyzer
    {

        /// <summary>
        /// Analyze bends by extracting bend line data from the FlatPattern "Bend-Lines" sketch.
        /// Matches the VBA CountBends logic exactly: traverses FlatPattern sub-features,
        /// finds the "Bend-Lines" ProfileFeature sketch, counts all sketch segments as bend count,
        /// and measures LINE segments for the longest bend line.
        /// Also detects bend direction (NeedsFlip) and max bend radius from OneBend features.
        /// </summary>
        public static BendInfo AnalyzeBends(IModelDoc2 model, bool countSuppressed = false)
        {
            var info = new BendInfo();
            if (model == null) return info;

            try
            {
                // === Pass 1: Extract bend count and longest bend line from "Bend-Lines" sketch ===
                // This matches VBA CountBends exactly.
                double longestLineMeters = 0.0;
                int sketchSegmentCount = 0;

                IFeature feat = model.FirstFeature() as IFeature;
                while (feat != null)
                {
                    string type = feat.GetTypeName2() ?? string.Empty;
                    if (type == "FlatPattern")
                    {
                        var subFeat = feat.GetFirstSubFeature() as IFeature;
                        while (subFeat != null)
                        {
                            string subType = subFeat.GetTypeName2() ?? string.Empty;
                            string subName = subFeat.Name ?? string.Empty;

                            // VBA: If SubFeat.GetTypeName2 = "ProfileFeature" And Left(SubFeat.Name, 10) = "Bend-Lines"
                            if (subType == "ProfileFeature" && subName.Length >= 10
                                && subName.Substring(0, 10) == "Bend-Lines")
                            {
                                ErrorHandler.DebugLog("[BendAnalyzer] " +
                                    string.Format("Found '{0}' (type={1}) under FlatPattern", subName, subType));

                                ISketch swSketch = subFeat.GetSpecificFeature2() as ISketch;
                                if (swSketch != null)
                                {
                                    object segObj = swSketch.GetSketchSegments();
                                    if (segObj is object[] vSketchSeg && vSketchSeg.Length > 0)
                                    {
                                        // VBA: reset to 0 for hidden parts to count last loop through
                                        longestLineMeters = 0.0;
                                        sketchSegmentCount = 0;

                                        ErrorHandler.DebugLog("[BendAnalyzer] " +
                                            string.Format("  Sketch has {0} segment(s)", vSketchSeg.Length));

                                        for (int i = 0; i < vSketchSeg.Length; i++)
                                        {
                                            var swSketchSeg = vSketchSeg[i] as ISketchSegment;
                                            if (swSketchSeg == null) continue;

                                            // VBA: If swSketchSeg.GetType = swSketchSegments_e.swSketchLINE Then
                                            try
                                            {
                                                int segType = swSketchSeg.GetType();
                                                double len = swSketchSeg.GetLength();
                                                ErrorHandler.DebugLog("[BendAnalyzer] " +
                                                    string.Format("  Seg[{0}]: type={1}, length={2:F6}m ({3:F4}in)",
                                                        i, segType, len, len * MetersToInches));

                                                if (segType == (int)swSketchSegments_e.swSketchLINE)
                                                {
                                                    if (len > longestLineMeters)
                                                        longestLineMeters = len;
                                                }
                                            }
                                            catch { /* swSketchSeg.GetType() can fail on corrupted segments */ }
                                        }
                                        // VBA: intFeatureCount = UBound(vSketchSeg) + 1
                                        sketchSegmentCount = vSketchSeg.Length;
                                    }
                                }
                            }
                            subFeat = subFeat.GetNextSubFeature() as IFeature;
                        }
                    }
                    feat = feat.GetNextFeature() as IFeature;
                }

                info.Count = sketchSegmentCount;
                // VBA: dblLongestLine = dblLongestLine * 39.36996  (convert meters to inches)
                info.LongestBendIn = longestLineMeters * MetersToInches;

                ErrorHandler.DebugLog("[BendAnalyzer] " +
                    string.Format("Result: Count={0}, LongestBendIn={1:F4}, longestLineMeters={2:F6}",
                        info.Count, info.LongestBendIn, longestLineMeters));

                // === Pass 2: Detect bend direction and max radius from OneBend features ===
                // VBA used a user prompt for flip detection; C# auto-detects from bend angles.
                // MaxRadiusIn is used by F325 roll forming calculation.
                bool hasUp = false, hasDown = false;
                feat = model.FirstFeature() as IFeature;
                while (feat != null)
                {
                    string type = feat.GetTypeName2() ?? string.Empty;
                    if (IsBendContainer(type))
                    {
                        var subFeat = feat.GetFirstSubFeature() as IFeature;
                        while (subFeat != null)
                        {
                            string subType = subFeat.GetTypeName2() ?? string.Empty;
                            if (subType == "OneBend")
                            {
                                ExtractBendProperties(subFeat, info, countSuppressed, ref hasUp, ref hasDown);
                            }
                            subFeat = subFeat.GetNextSubFeature() as IFeature;
                        }
                    }
                    else if (feat.GetTypeName2() == "OneBend")
                    {
                        ExtractBendProperties(feat, info, countSuppressed, ref hasUp, ref hasDown);
                    }
                    feat = feat.GetNextFeature() as IFeature;
                }
                info.NeedsFlip = hasUp && hasDown;
            }
            catch { }

            return info;
        }

        // Feature types that contain bend sub-features (from InsertBends2)
        private static readonly string[] BendContainerTypes = { "ProcessBends", "FlattenBends" };

        private static bool IsBendContainer(string typeName)
        {
            for (int i = 0; i < BendContainerTypes.Length; i++)
            {
                if (typeName == BendContainerTypes[i]) return true;
            }
            return false;
        }

        /// <summary>
        /// Extract bend direction and max radius from an OneBend feature.
        /// Does NOT modify info.Count or info.LongestBendIn — those come from the sketch.
        /// </summary>
        private static void ExtractBendProperties(IFeature feat, BendInfo info, bool countSuppressed, ref bool hasUp, ref bool hasDown)
        {
            bool suppressed = false;
            try { suppressed = feat.IsSuppressed(); } catch { }
            if (suppressed && !countSuppressed) return;

            try
            {
                object defObj = feat.GetDefinition();
                if (defObj is IOneBendFeatureData one)
                {
                    double rIn = one.BendRadius * MetersToInches;
                    if (rIn > info.MaxRadiusIn) info.MaxRadiusIn = rIn;

                    double ang = 0.0;
                    try { ang = one.BendAngle; } catch { }
                    if (ang > 0) hasUp = true; else if (ang < 0) hasDown = true;
                }
            }
            catch { }
        }
    }
}
