using System;
using System.Collections.Generic;
using System.Linq;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.SheetMetal
{
    /// <summary>
    /// Colors sheet metal faces to indicate top (good/cosmetic) vs bottom (bad) side.
    /// Green = top/fold-up side (#3 finish), Red = bottom/machine side.
    /// Toggle on/off via toolbar button.
    /// </summary>
    public sealed class SideIndicatorService
    {
        // Track which models currently have indicators applied (by full path)
        private readonly HashSet<string> _activeModels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Green for good/top side
        private static readonly double[] TopColor = { 0.0, 0.8, 0.0, 1.0, 1.0, 0.3, 0.4, 0.0, 0.0 };
        // Red for bad/bottom side
        private static readonly double[] BottomColor = { 0.8, 0.0, 0.0, 1.0, 1.0, 0.3, 0.4, 0.0, 0.0 };
        // Light gray for edge/bend faces
        private static readonly double[] EdgeColor = { 0.7, 0.7, 0.7, 1.0, 1.0, 0.3, 0.4, 0.0, 0.0 };

        private const double NormalTolerance = 0.85; // dot product threshold for same/opposite direction

        /// <summary>
        /// Toggles side indicators on the active document.
        /// First call applies colors, second call removes them.
        /// </summary>
        public void Toggle(ISldWorks swApp)
        {
            const string proc = "SideIndicator.Toggle";
            ErrorHandler.PushCallStack(proc);
            try
            {
                var model = swApp.ActiveDoc as IModelDoc2;
                if (model == null)
                {
                    swApp.SendMsgToUser("No document is open.");
                    return;
                }

                string path = model.GetPathName() ?? model.GetTitle();
                int docType = model.GetType();

                if (_activeModels.Contains(path))
                {
                    // Remove indicators
                    if (docType == (int)swDocumentTypes_e.swDocASSEMBLY)
                        RemoveFromAssembly(model);
                    else
                        RemoveFromPart(model);

                    _activeModels.Remove(path);
                    model.GraphicsRedraw2();
                    ErrorHandler.DebugLog($"{proc}: Removed side indicators from {path}");
                }
                else
                {
                    // Apply indicators
                    int count = 0;
                    if (docType == (int)swDocumentTypes_e.swDocASSEMBLY)
                        count = ApplyToAssembly(swApp, model);
                    else
                        count = ApplyToPart(model);

                    if (count > 0)
                    {
                        _activeModels.Add(path);
                        model.GraphicsRedraw2();
                        ErrorHandler.DebugLog($"{proc}: Applied side indicators to {count} bodies in {path}");
                    }
                    else
                    {
                        swApp.SendMsgToUser("No sheet metal bodies found to color.");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex, ErrorHandler.LogLevel.Error);
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Returns true if side indicators are currently active on the given model.
        /// </summary>
        public bool IsActive(IModelDoc2 model)
        {
            if (model == null) return false;
            string path = model.GetPathName() ?? model.GetTitle();
            return _activeModels.Contains(path);
        }

        #region Apply

        private int ApplyToPart(IModelDoc2 model)
        {
            var part = model as IPartDoc;
            if (part == null) return 0;

            var bodiesRaw = part.GetBodies2((int)swBodyType_e.swSolidBody, true);
            if (bodiesRaw == null) return 0;
            var bodies = ((object[])bodiesRaw).Cast<IBody2>().ToList();

            double[] topNormal = GetFixedFaceNormal(model);
            if (topNormal == null)
            {
                ErrorHandler.DebugLog("[SideIndicator] Could not determine fixed face normal, falling back to largest face");
                topNormal = GetLargestPlanarFaceNormal(bodies);
            }
            if (topNormal == null) return 0;

            int count = 0;
            foreach (var body in bodies)
            {
                if (ApplyToBody(body, topNormal))
                    count++;
            }
            return count;
        }

        private int ApplyToAssembly(ISldWorks swApp, IModelDoc2 model)
        {
            var config = model.ConfigurationManager.ActiveConfiguration;
            if (config == null) return 0;

            var rootComp = config.GetRootComponent3(true) as IComponent2;
            if (rootComp == null) return 0;

            int count = 0;
            ApplyToComponentTree(rootComp, ref count);
            return count;
        }

        private void ApplyToComponentTree(IComponent2 comp, ref int count)
        {
            var childrenRaw = comp.GetChildren() as object[];
            if (childrenRaw == null) return;

            foreach (var childObj in childrenRaw)
            {
                var child = childObj as IComponent2;
                if (child == null) continue;
                if (child.IsSuppressed()) continue;

                var compDoc = child.GetModelDoc2() as IModelDoc2;
                if (compDoc == null) continue;

                int docType = compDoc.GetType();
                if (docType == (int)swDocumentTypes_e.swDocPART)
                {
                    // Only color sheet metal parts
                    if (SwGeometryHelper.HasSheetMetalFeature(compDoc))
                    {
                        var part = compDoc as IPartDoc;
                        if (part != null)
                        {
                            var bodiesRaw = part.GetBodies2((int)swBodyType_e.swSolidBody, true);
                            if (bodiesRaw != null)
                            {
                                var bodies = ((object[])bodiesRaw).Cast<IBody2>().ToList();
                                double[] topNormal = GetFixedFaceNormal(compDoc);
                                if (topNormal == null)
                                    topNormal = GetLargestPlanarFaceNormal(bodies);

                                if (topNormal != null)
                                {
                                    foreach (var body in bodies)
                                    {
                                        if (ApplyToBody(body, topNormal))
                                            count++;
                                    }
                                }
                            }
                        }
                    }
                }
                else if (docType == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    // Recurse into sub-assemblies
                    ApplyToComponentTree(child, ref count);
                }
            }
        }

        private bool ApplyToBody(IBody2 body, double[] topNormal)
        {
            var facesRaw = body.GetFaces();
            if (facesRaw == null) return false;

            var faces = ((object[])facesRaw).Cast<IFace2>().ToList();
            if (faces.Count == 0) return false;

            foreach (var face in faces)
            {
                try
                {
                    var surface = face.GetSurface() as ISurface;
                    if (surface == null) continue;

                    if (surface.IsPlane())
                    {
                        double[] faceNormal = face.Normal as double[];
                        if (faceNormal == null || faceNormal.Length < 3) continue;

                        double dot = faceNormal[0] * topNormal[0]
                                   + faceNormal[1] * topNormal[1]
                                   + faceNormal[2] * topNormal[2];

                        if (dot > NormalTolerance)
                            face.MaterialPropertyValues = (double[])TopColor.Clone();
                        else if (dot < -NormalTolerance)
                            face.MaterialPropertyValues = (double[])BottomColor.Clone();
                        else
                            face.MaterialPropertyValues = (double[])EdgeColor.Clone();
                    }
                    else
                    {
                        // Cylindrical (bends) and other non-planar faces
                        face.MaterialPropertyValues = (double[])EdgeColor.Clone();
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandler.DebugLog($"[SideIndicator] Error coloring face: {ex.Message}");
                }
            }
            return true;
        }

        #endregion

        #region Remove

        private void RemoveFromPart(IModelDoc2 model)
        {
            var part = model as IPartDoc;
            if (part == null) return;

            var bodiesRaw = part.GetBodies2((int)swBodyType_e.swSolidBody, true);
            if (bodiesRaw == null) return;

            foreach (var bodyObj in (object[])bodiesRaw)
            {
                var body = bodyObj as IBody2;
                if (body == null) continue;
                ClearBodyColors(body);
            }
        }

        private void RemoveFromAssembly(IModelDoc2 model)
        {
            var config = model.ConfigurationManager.ActiveConfiguration;
            if (config == null) return;

            var rootComp = config.GetRootComponent3(true) as IComponent2;
            if (rootComp == null) return;

            RemoveFromComponentTree(rootComp);
        }

        private void RemoveFromComponentTree(IComponent2 comp)
        {
            var childrenRaw = comp.GetChildren() as object[];
            if (childrenRaw == null) return;

            foreach (var childObj in childrenRaw)
            {
                var child = childObj as IComponent2;
                if (child == null) continue;
                if (child.IsSuppressed()) continue;

                var compDoc = child.GetModelDoc2() as IModelDoc2;
                if (compDoc == null) continue;

                int docType = compDoc.GetType();
                if (docType == (int)swDocumentTypes_e.swDocPART)
                {
                    if (SwGeometryHelper.HasSheetMetalFeature(compDoc))
                    {
                        var part = compDoc as IPartDoc;
                        if (part == null) continue;
                        var bodiesRaw = part.GetBodies2((int)swBodyType_e.swSolidBody, true);
                        if (bodiesRaw == null) continue;
                        foreach (var bodyObj in (object[])bodiesRaw)
                        {
                            var body = bodyObj as IBody2;
                            if (body != null) ClearBodyColors(body);
                        }
                    }
                }
                else if (docType == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    RemoveFromComponentTree(child);
                }
            }
        }

        private void ClearBodyColors(IBody2 body)
        {
            // Clear face-level overrides
            var facesRaw = body.GetFaces();
            if (facesRaw != null)
            {
                foreach (var faceObj in (object[])facesRaw)
                {
                    var face = faceObj as IFace2;
                    if (face == null) continue;
                    try { face.MaterialPropertyValues = null; }
                    catch { }
                }
            }

            // Clear body-level override
            try
            {
                body.RemoveMaterialProperty(
                    (int)swInConfigurationOpts_e.swAllConfiguration, null);
            }
            catch { }
        }

        #endregion

        #region Fixed Face Detection

        /// <summary>
        /// Gets the outward normal of the fixed face from the FlatPattern feature.
        /// This is the definitive "top" / "cosmetic" / "fold up" side.
        /// </summary>
        private double[] GetFixedFaceNormal(IModelDoc2 model)
        {
            try
            {
                IFeature flatFeat = FindFlatPatternFeature(model);
                if (flatFeat == null) return null;

                var flatData = flatFeat.GetDefinition() as FlatPatternFeatureData;
                if (flatData == null) return null;

                bool accessed = flatData.AccessSelections(model, null);
                if (!accessed)
                {
                    ErrorHandler.DebugLog("[SideIndicator] Could not access FlatPattern selections");
                    return null;
                }

                try
                {
                    var fixedFace = flatData.FixedFace2 as IFace2;
                    if (fixedFace == null) return null;

                    var normal = fixedFace.Normal as double[];
                    if (normal == null || normal.Length < 3) return null;

                    return new double[] { normal[0], normal[1], normal[2] };
                }
                finally
                {
                    flatData.ReleaseSelectionAccess();
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[SideIndicator] Error getting fixed face: {ex.Message}");
                return null;
            }
        }

        private IFeature FindFlatPatternFeature(IModelDoc2 model)
        {
            if (model == null) return null;

            var feat = model.FirstFeature() as IFeature;
            while (feat != null)
            {
                string type = feat.GetTypeName2() ?? string.Empty;
                if (type.Equals("FlatPattern", StringComparison.OrdinalIgnoreCase))
                    return feat;
                feat = feat.GetNextFeature() as IFeature;
            }
            return null;
        }

        /// <summary>
        /// Fallback: use the normal of the largest planar face as the "top" direction.
        /// Used when FlatPattern feature is not available (e.g. imported sheet metal).
        /// </summary>
        private double[] GetLargestPlanarFaceNormal(List<IBody2> bodies)
        {
            IFace2 largest = null;
            double maxArea = 0;

            foreach (var body in bodies)
            {
                var facesRaw = body.GetFaces();
                if (facesRaw == null) continue;

                foreach (var faceObj in (object[])facesRaw)
                {
                    var face = faceObj as IFace2;
                    if (face == null) continue;

                    try
                    {
                        var surface = face.GetSurface() as ISurface;
                        if (surface == null || !surface.IsPlane()) continue;

                        double area = face.GetArea();
                        if (area > maxArea)
                        {
                            maxArea = area;
                            largest = face;
                        }
                    }
                    catch { }
                }
            }

            if (largest == null) return null;
            var normal = largest.Normal as double[];
            if (normal == null || normal.Length < 3) return null;
            return new double[] { normal[0], normal[1], normal[2] };
        }

        #endregion
    }
}
