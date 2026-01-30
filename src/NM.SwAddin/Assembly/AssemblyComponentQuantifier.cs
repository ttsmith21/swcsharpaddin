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
            /// <summary>
            /// Immediate parent assembly path (for multi-level BOM support).
            /// </summary>
            public string ParentAssemblyPath { get; set; }
            public string Key => BuildKey(FilePath, Configuration);
        }

        /// <summary>
        /// Represents a node in the assembly hierarchy for multi-level BOM export.
        /// </summary>
        public sealed class BomNode
        {
            public string FilePath { get; set; }
            public string Configuration { get; set; }
            public string PartNumber { get; set; }
            public int Quantity { get; set; }
            public bool IsAssembly { get; set; }
            public List<BomNode> Children { get; } = new List<BomNode>();
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
        /// Includes all levels; skips suppressed components. Tracks immediate parent assembly.
        /// </summary>
        public Dictionary<string, ComponentQuantity> CollectViaRecursion(IAssemblyDoc asm)
        {
            var result = new Dictionary<string, ComponentQuantity>(StringComparer.OrdinalIgnoreCase);
            var model = (IModelDoc2)asm;
            var cfg = model?.ConfigurationManager?.ActiveConfiguration;
            var root = cfg?.GetRootComponent3(true);
            if (root == null) return result;

            // Get root assembly path as the initial parent
            string rootPath = Safe(() => model.GetPathName()) ?? string.Empty;
            Traverse(root, result, rootPath);
            return result;
        }

        /// <summary>
        /// Collects the full assembly hierarchy as a tree of BomNodes.
        /// For multi-level assembly support with deep hierarchies.
        /// </summary>
        public BomNode CollectHierarchy(IAssemblyDoc asm)
        {
            var model = (IModelDoc2)asm;
            if (model == null) return null;

            var cfg = model.ConfigurationManager?.ActiveConfiguration;
            var root = cfg?.GetRootComponent3(true);
            if (root == null) return null;

            var rootNode = new BomNode
            {
                FilePath = Safe(() => model.GetPathName()) ?? "",
                Configuration = Safe(() => cfg.Name) ?? "",
                PartNumber = System.IO.Path.GetFileNameWithoutExtension(model.GetPathName() ?? ""),
                IsAssembly = true,
                Quantity = 1
            };

            // Collect children
            var kids = root.GetChildren();
            if (kids is Array arr)
            {
                CollectChildrenHierarchy(arr, rootNode);
            }

            return rootNode;
        }

        private void CollectChildrenHierarchy(Array children, BomNode parent)
        {
            // Group by path+config to aggregate quantities
            var grouped = new Dictionary<string, BomNode>(StringComparer.OrdinalIgnoreCase);

            foreach (var o in children)
            {
                var comp = o as IComponent2;
                if (comp == null) continue;

                try
                {
                    int sup = comp.GetSuppression2();
                    if (sup == (int)swComponentSuppressionState_e.swComponentSuppressed)
                        continue;

                    string path = Safe(() => comp.GetPathName()) ?? string.Empty;
                    string cfg = Safe(() => comp.ReferencedConfiguration) ?? string.Empty;
                    string key = BuildKey(path, cfg);
                    var childDoc = comp.GetModelDoc2() as IModelDoc2;
                    bool isAsm = childDoc != null && childDoc.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY;

                    if (!grouped.TryGetValue(key, out var node))
                    {
                        node = new BomNode
                        {
                            FilePath = path,
                            Configuration = cfg,
                            PartNumber = System.IO.Path.GetFileNameWithoutExtension(path),
                            IsAssembly = isAsm,
                            Quantity = 0
                        };
                        grouped[key] = node;
                        parent.Children.Add(node);
                    }
                    node.Quantity++;

                    // Recurse into sub-assemblies
                    if (isAsm)
                    {
                        var subKids = comp.GetChildren();
                        if (subKids is Array subArr && subArr.Length > 0)
                        {
                            CollectChildrenHierarchy(subArr, node);
                        }
                    }
                }
                catch { }
            }
        }

        private void Traverse(IComponent2 comp, Dictionary<string, ComponentQuantity> counts, string parentAssemblyPath)
        {
            if (comp == null) return;
            try
            {
                int sup = comp.GetSuppression2();
                if (sup == (int)swComponentSuppressionState_e.swComponentSuppressed)
                    return;

                var childDoc = comp.GetModelDoc2() as IModelDoc2;
                string compPath = Safe(() => comp.GetPathName()) ?? string.Empty;

                // Count parts only
                if (childDoc != null && childDoc.GetType() == (int)swDocumentTypes_e.swDocPART)
                {
                    string cfg = Safe(() => comp.ReferencedConfiguration) ?? string.Empty;
                    string key = BuildKey(compPath, cfg);
                    if (!counts.TryGetValue(key, out var q))
                    {
                        q = new ComponentQuantity
                        {
                            FilePath = compPath,
                            Configuration = cfg,
                            Quantity = 0,
                            ParentAssemblyPath = parentAssemblyPath
                        };
                        counts[key] = q;
                    }
                    q.Quantity++;
                }

                // Recurse into sub-assemblies (using sub-assembly path as new parent)
                if (childDoc != null && childDoc.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    var kids = comp.GetChildren();
                    if (kids is Array arr)
                    {
                        foreach (var o in arr)
                        {
                            // Sub-assembly becomes the parent for its children
                            Traverse(o as IComponent2, counts, compPath);
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
