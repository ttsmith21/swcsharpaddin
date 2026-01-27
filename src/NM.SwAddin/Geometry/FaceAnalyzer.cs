using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Geometry
{
    public sealed class FaceAnalyzer
    {
        // Gets the currently selected face or falls back to fixed/largest face
        public IFace2 GetProcessingFace(IModelDoc2 model)
        {
            if (model == null) return null;
            var sel = GetSelectedFace(model);
            if (sel != null) return sel;

            var fixedFace = GetFixedFace(model);
            if (fixedFace != null) return fixedFace;

            return GetLargestFace(model);
        }

        public IFace2 GetSelectedFace(IModelDoc2 model)
        {
            try
            {
                var selMgr = model.SelectionManager as ISelectionMgr;
                if (selMgr == null) return null;
                int count = selMgr.GetSelectedObjectCount2(-1);
                for (int i = 1; i <= count; i++)
                {
                    var type = (swSelectType_e)selMgr.GetSelectedObjectType3(i, -1);
                    if (type == swSelectType_e.swSelFACES)
                    {
                        return selMgr.GetSelectedObject6(i, -1) as IFace2;
                    }
                }
            }
            catch { }
            return null;
        }

        // Best-effort fixed face: prefer largest planar face on main body for SM parts
        public IFace2 GetFixedFace(IModelDoc2 model)
        {
            try
            {
                if (!NM.SwAddin.SolidWorksApiWrapper.HasSheetMetalFeature(model)) return null;
                var body = NM.SwAddin.SolidWorksApiWrapper.GetMainBody(model);
                if (body == null) return null;
                return NM.SwAddin.SolidWorksApiWrapper.GetLargestPlanarFace(body);
            }
            catch { return null; }
        }

        public IFace2 GetLargestFace(IModelDoc2 model)
        {
            try
            {
                var part = model as IPartDoc;
                if (part == null) return null;
                var bodiesObj = part.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
                if (bodiesObj == null || bodiesObj.Length == 0) return null;
                IFace2 largest = null;
                double maxArea = 0;
                foreach (var bo in bodiesObj)
                {
                    var body = bo as IBody2; if (body == null) continue;
                    var faces = body.GetFaces() as object[]; if (faces == null) continue;
                    foreach (var fo in faces)
                    {
                        var f = fo as IFace2; if (f == null) continue;
                        try
                        {
                            double area = 0;
                            try { area = f.GetArea(); } catch { area = 0; }
                            if (area > maxArea)
                            {
                                maxArea = area; largest = f;
                            }
                        }
                        catch { }
                    }
                }
                return largest;
            }
            catch { return null; }
        }
    }
}
