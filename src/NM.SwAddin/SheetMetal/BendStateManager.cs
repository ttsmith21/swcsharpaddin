using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.SheetMetal
{
    public sealed class BendStateManager
    {
        public bool SelectFlatPattern(IModelDoc2 model, out IConfiguration flatConfig)
        {
            flatConfig = null;
            if (model == null) return false;

            var names = model.GetConfigurationNames() as string[];
            string flatName = null;
            if (names != null)
            {
                foreach (var n in names)
                {
                    if (string.IsNullOrWhiteSpace(n)) continue;
                    var u = n.ToUpperInvariant();
                    if (u.Contains("FLAT") || u.Contains("FLATPATTERN")) { flatName = n; break; }
                }
            }

            if (string.IsNullOrEmpty(flatName))
            {
                try { ActivateFlatPatternFeature(model); } catch { }
                names = model.GetConfigurationNames() as string[];
                if (names != null)
                {
                    foreach (var n in names)
                    {
                        if (string.IsNullOrWhiteSpace(n)) continue;
                        var u = n.ToUpperInvariant();
                        if (u.Contains("FLAT") || u.Contains("FLATPATTERN")) { flatName = n; break; }
                    }
                }
            }

            if (string.IsNullOrEmpty(flatName)) return false;

            try
            {
                model.ShowConfiguration2(flatName);
                flatConfig = model.GetConfigurationByName(flatName) as IConfiguration;
                try { model.EditRebuild3(); }
                catch { try { model.ForceRebuild3(false); } catch { } }
                return true;
            }
            catch { return false; }
        }

        // Toggle/create flat pattern by finding FlatPattern feature and unsuppressing it
        private void ActivateFlatPatternFeature(IModelDoc2 model)
        {
            IFeature feat = model.FirstFeature() as IFeature;
            while (feat != null)
            {
                var type = feat.GetTypeName2() ?? string.Empty;
                if (type.IndexOf("FlatPattern", System.StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    try { feat.SetSuppression2(1, 2, null); } catch { }
                    break;
                }
                feat = feat.GetNextFeature() as IFeature;
            }
        }
    }
}
