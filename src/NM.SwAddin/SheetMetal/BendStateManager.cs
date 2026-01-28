using System;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.SheetMetal
{
    /// <summary>
    /// Manages sheet metal bend state transitions (flattened vs folded).
    /// Ported from VBA SP.bas UnsuppressFlatten/SuppressFlatten functions.
    /// </summary>
    public sealed class BendStateManager
    {
        /// <summary>
        /// Sets part to flattened state by unsuppressing the FlatPattern feature.
        /// </summary>
        public bool FlattenPart(IModelDoc2 model)
        {
            const string proc = nameof(FlattenPart);
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (model == null) return false;

                var flatFeat = FindFlatPatternFeature(model);
                if (flatFeat == null)
                {
                    ErrorHandler.DebugLog($"{proc}: No FlatPattern feature found");
                    return false;
                }

                // Unsuppress the flat pattern feature (swUnSuppressFeature = 1, swAllConfiguration = 2)
                bool result = flatFeat.SetSuppression2(
                    (int)swFeatureSuppressionAction_e.swUnSuppressFeature,
                    (int)swInConfigurationOpts_e.swAllConfiguration,
                    null);

                if (result)
                {
                    try { model.EditRebuild3(); }
                    catch { try { model.ForceRebuild3(false); } catch { } }
                }

                return result;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                return false;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        /// <summary>
        /// Sets part to folded/bent state by suppressing the FlatPattern feature.
        /// Ported from VBA SP.bas SuppressFlatten.
        /// </summary>
        public bool UnFlattenPart(IModelDoc2 model)
        {
            const string proc = nameof(UnFlattenPart);
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (model == null) return false;

                var flatFeat = FindFlatPatternFeature(model);
                if (flatFeat == null)
                {
                    // No flat pattern = already folded
                    return true;
                }

                // Suppress the flat pattern feature (swSuppressFeature = 0, swAllConfiguration = 2)
                bool result = flatFeat.SetSuppression2(
                    (int)swFeatureSuppressionAction_e.swSuppressFeature,
                    (int)swInConfigurationOpts_e.swAllConfiguration,
                    null);

                if (result)
                {
                    try { model.EditRebuild3(); }
                    catch { try { model.ForceRebuild3(false); } catch { } }
                }

                return result;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                return false;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        /// <summary>
        /// Checks if the part is currently flattened (FlatPattern unsuppressed).
        /// </summary>
        public bool IsFlattened(IModelDoc2 model)
        {
            if (model == null) return false;

            var flatFeat = FindFlatPatternFeature(model);
            if (flatFeat == null) return false;

            // Check if feature is NOT suppressed (meaning flat pattern is active)
            try
            {
                // IsSuppressed returns true if the feature is suppressed
                return !flatFeat.IsSuppressed();
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Finds the FlatPattern feature in the model.
        /// </summary>
        public IFeature FindFlatPatternFeature(IModelDoc2 model)
        {
            if (model == null) return null;

            IFeature feat = model.FirstFeature() as IFeature;
            while (feat != null)
            {
                string type = feat.GetTypeName2() ?? string.Empty;
                if (type.Equals("FlatPattern", StringComparison.OrdinalIgnoreCase))
                {
                    return feat;
                }
                feat = feat.GetNextFeature() as IFeature;
            }
            return null;
        }

        /// <summary>
        /// Finds and selects the Sheet-Metal feature for modification.
        /// </summary>
        public IFeature FindSheetMetalFeature(IModelDoc2 model)
        {
            if (model == null) return null;

            IFeature feat = model.FirstFeature() as IFeature;
            while (feat != null)
            {
                string type = feat.GetTypeName2() ?? string.Empty;
                // Sheet metal features: "SheetMetal", "SMBaseFlange", "EdgeFlange", etc.
                if (type.Equals("SheetMetal", StringComparison.OrdinalIgnoreCase) ||
                    type.Equals("SMBaseFlange", StringComparison.OrdinalIgnoreCase))
                {
                    return feat;
                }
                feat = feat.GetNextFeature() as IFeature;
            }
            return null;
        }

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
