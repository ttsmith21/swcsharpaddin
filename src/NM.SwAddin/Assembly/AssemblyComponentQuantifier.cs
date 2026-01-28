using System;
using System.Collections.Generic;
using System.Linq;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.AssemblyProcessing
{
    /// <summary>
    /// Hybrid component quantifier:
    /// - Primary: BOM table counts (most accurate for ERP quantities)
    /// - Cross-check: recursive traversal counts
    /// Returns BOM as authoritative and logs any mismatches.
    /// </summary>
    public sealed class AssemblyComponentQuantifier
    {
        public sealed class ComponentQuantity
        {
            public string FilePath { get; set; }
            public string Configuration { get; set; }
            public int Quantity { get; set; }
            public string Key => BuildKey(FilePath, Configuration);
        }

        public static string BuildKey(string path, string cfg)
            => ($"{(path ?? string.Empty).ToLowerInvariant()}::{cfg ?? string.Empty}");

        /// <summary>
        /// Collect quantities using BOM (authoritative) plus recursive cross-check.
        /// </summary>
        public Dictionary<string, ComponentQuantity> CollectQuantitiesHybrid(IAssemblyDoc asm, string bomTemplatePath = "")
        {
            if (asm == null) return new Dictionary<string, ComponentQuantity>(StringComparer.OrdinalIgnoreCase);
            var model = (IModelDoc2)asm;

            // 1) Try BOM-based
            var bomCounts = TryCollectViaBom(asm, bomTemplatePath);

            // 2) Cross-check via recursion
            var recCounts = CollectViaRecursion(asm);

            // 3) Reconcile: prefer BOM, log discrepancies
            foreach (var kv in recCounts)
            {
                if (bomCounts.TryGetValue(kv.Key, out var bomItem))
                {
                    if (bomItem.Quantity != kv.Value.Quantity)
                    {
                        ErrorHandler.HandleError(nameof(AssemblyComponentQuantifier),
                            $"Quantity mismatch for {kv.Key}: BOM={bomItem.Quantity}, Rec={kv.Value.Quantity}", null, ErrorHandler.LogLevel.Warning);
                    }
                }
                else
                {
                    // Not present in BOM, but found by recursion
                    ErrorHandler.HandleError(nameof(AssemblyComponentQuantifier),
                        $"Recursion found component not in BOM: {kv.Key} x{kv.Value.Quantity}", null, ErrorHandler.LogLevel.Warning);
                }
            }

            // If BOM is empty (failed), fall back to recursion
            if (bomCounts.Count == 0 && recCounts.Count > 0)
            {
                ErrorHandler.HandleError(nameof(AssemblyComponentQuantifier), "BOM collection failed/empty. Falling back to recursive quantities.", null, ErrorHandler.LogLevel.Warning);
                return recCounts;
            }
            return bomCounts;
        }

        /// <summary>
        /// Recursively counts unique part/config occurrences in the assembly tree.
        /// Includes all levels; skips suppressed components.
        /// </summary>
        public Dictionary<string, ComponentQuantity> CollectViaRecursion(IAssemblyDoc asm)
        {
            var result = new Dictionary<string, ComponentQuantity>(StringComparer.OrdinalIgnoreCase);
            var model = (IModelDoc2)asm;
            var cfg = model?.ConfigurationManager?.ActiveConfiguration;
            var root = cfg?.GetRootComponent3(true);
            if (root == null) return result;

            Traverse(root, result);
            return result;
        }

        private void Traverse(IComponent2 comp, Dictionary<string, ComponentQuantity> counts)
        {
            if (comp == null) return;
            try
            {
                int sup = comp.GetSuppression2();
                if (sup == (int)swComponentSuppressionState_e.swComponentSuppressed)
                    return;

                var childDoc = comp.GetModelDoc2() as IModelDoc2;
                // Count parts only
                if (childDoc != null && childDoc.GetType() == (int)swDocumentTypes_e.swDocPART)
                {
                    string path = Safe(() => comp.GetPathName()) ?? string.Empty;
                    string cfg = Safe(() => comp.ReferencedConfiguration) ?? string.Empty;
                    string key = BuildKey(path, cfg);
                    if (!counts.TryGetValue(key, out var q))
                    {
                        q = new ComponentQuantity { FilePath = path, Configuration = cfg, Quantity = 0 };
                        counts[key] = q;
                    }
                    q.Quantity++;
                }

                // Recurse into sub-assemblies
                if (childDoc != null && childDoc.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    var kids = comp.GetChildren();
                    if (kids is Array arr)
                    {
                        foreach (var o in arr)
                        {
                            Traverse(o as IComponent2, counts);
                        }
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Use a temporary BOM table to read unique parts and their counts.
        /// Cleans up the inserted BOM feature afterward.
        /// </summary>
        public Dictionary<string, ComponentQuantity> TryCollectViaBom(IAssemblyDoc asm, string bomTemplatePath = "")
        {
            var result = new Dictionary<string, ComponentQuantity>(StringComparer.OrdinalIgnoreCase);
            ITableAnnotation table = null;
            IBomTableAnnotation bom = null;
            IBomFeature bomFeat = null;
            IFeature bomFeature = null;
            var model = (IModelDoc2)asm;

            try
            {
                var ext = model.Extension;
                var activeCfg = Safe(() => model.ConfigurationManager?.ActiveConfiguration?.Name) ?? "";
                // Insert BOM; let SW use default template if path is empty
                // InsertBomTable3 signature may vary by SW version; use positional params
                bom = ext.InsertBomTable3(
                    bomTemplatePath ?? string.Empty,
                    0,  // X position
                    0,  // Y position
                    (int)swBomType_e.swBomType_PartsOnly,
                    activeCfg,
                    false, // displayQuantity
                    (int)swNumberingType_e.swIndentedBOMNotSet,
                    false  // showDetailedCutList
                ) as IBomTableAnnotation;

                if (bom == null)
                {
                    ErrorHandler.HandleError(nameof(AssemblyComponentQuantifier), "InsertBomTable3 returned null.", null, ErrorHandler.LogLevel.Warning);
                    return result;
                }

                bomFeat = bom.BomFeature;
                table = bom as ITableAnnotation;
                // Force rebuild so BOM populates
                try { model.ForceRebuild3(true); } catch { }

                int rows = 0;
                try { rows = table.RowCount; } catch { rows = 0; }
                // Data rows start at 1; row 0 is header
                for (int r = 1; r < rows; r++)
                {
                    // Count components in this row
                    int count = 0;
                    string itemNo = null; string modelName = null;
                    try { count = bom.GetComponentsCount2(r, "", out itemNo, out modelName); } catch { count = 0; }

                    // Try to get first component to derive file path + config
                    string path = string.Empty; string cfgName = string.Empty;
                    object compsObj = null;
                    try
                    {
                        // Use GetComponents2 with row index and configuration
                        try { compsObj = bom.GetComponents2(r, activeCfg); } catch { compsObj = null; }

                        if (compsObj is object[] arr && arr.Length > 0)
                        {
                            var first = arr[0] as IComponent2;
                            if (first != null)
                            {
                                path = Safe(() => first.GetPathName()) ?? string.Empty;
                                cfgName = Safe(() => first.ReferencedConfiguration) ?? string.Empty;
                            }
                        }
                    }
                    catch { }

                    // Fallback: if no component object, try using modelName (may be filename)
                    if (string.IsNullOrWhiteSpace(path))
                    {
                        path = modelName ?? string.Empty;
                    }

                    string key = BuildKey(path, cfgName);
                    if (!result.TryGetValue(key, out var q))
                    {
                        q = new ComponentQuantity { FilePath = path, Configuration = cfgName, Quantity = 0 };
                        result[key] = q;
                    }
                    q.Quantity += Math.Max(0, count);
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(nameof(AssemblyComponentQuantifier), "BOM collection error", ex, ErrorHandler.LogLevel.Warning);
            }
            finally
            {
                // Cleanup BOM feature
                try
                {
                    if (bomFeat != null)
                    {
                        bomFeature = bomFeat.GetFeature();
                    }
                }
                catch { bomFeature = null; }

                try
                {
                    if (bomFeature != null)
                    {
                        bomFeature.Select2(false, 0);
                        // swDelete_Children = 1 (deletes without prompt)
                        model.Extension.DeleteSelection2(1);
                        try { model.ForceRebuild3(true); } catch { }
                    }
                }
                catch { }
            }

            return result;
        }

        private static T Safe<T>(Func<T> f)
        {
            try { return f(); } catch { return default(T); }
        }
    }
}
