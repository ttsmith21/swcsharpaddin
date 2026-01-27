using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace swcsharpaddin
{
    internal class StepFileCleanupService
    {
        private readonly ISldWorks _swApp;
        private const int DUPLICATE_GROUP_THRESHOLD = 5; // Minimum number of duplicates to group into a folder

        // Reorder position constants (match VBA values)
        private const int REORDER_AFTER = 1;
        private const int REORDER_BEFORE = 2;
        private const int REORDER_LAST_IN_FOLDER = 3;
        private const int REORDER_FIRST_IN_FOLDER = 4;

        public StepFileCleanupService(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        public void Run()
        {
            // 1) Validate active doc is an assembly
            var model = _swApp?.ActiveDoc as IModelDoc2;
            if (model == null)
            {
                System.Windows.Forms.MessageBox.Show("No active document found.");
                return;
            }

            var asm = model as IAssemblyDoc;
            if (asm == null)
            {
                System.Windows.Forms.MessageBox.Show("Active document is not an assembly.");
                return;
            }

            // 2) Build unique assembly list (root + subassemblies)
            var assemblies = BuildUniqueAssemblyList(asm);

            // 3) Process all assemblies (reorder and group duplicates)
            ProcessAllAssemblies(assemblies);

            // 4) Rebuild
            try { model.ForceRebuild3(true); } catch { }
        }

        private List<IComponent2> BuildUniqueAssemblyList(IAssemblyDoc asmDoc)
        {
            var unique = new List<IComponent2>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var model = (IModelDoc2)asmDoc;
            var cfg = model?.ConfigurationManager?.ActiveConfiguration;
            var root = cfg?.GetRootComponent3(true);
            if (root == null)
                return unique;

            if (!ResolveComponent(root))
            {
                System.Diagnostics.Debug.WriteLine("Failed to resolve root component");
                return unique;
            }

            AddAssemblyToList(root, unique, seen);
            TraverseAssemblies(root, unique, seen);
            return unique;
        }

        private void AddAssemblyToList(IComponent2 comp, List<IComponent2> list, HashSet<string> seen)
        {
            if (comp == null) return;
            // Ensure resolved if possible
            if (!ResolveComponent(comp)) return;

            var modelDoc = comp.GetModelDoc2() as IModelDoc2;
            string key;
            if (modelDoc != null)
            {
                var cfg = modelDoc.ConfigurationManager?.ActiveConfiguration;
                key = ($"{Safe(modelDoc.GetTitle())}_{Safe(cfg?.Name)}");
            }
            else
            {
                key = Safe(comp.Name2);
            }

            if (seen.Add(key))
            {
                list.Add(comp);
            }
        }

        private void TraverseAssemblies(IComponent2 comp, List<IComponent2> list, HashSet<string> seen)
        {
            if (comp == null) return;
            var childrenObj = comp.GetChildren();
            if (childrenObj is Array arr)
            {
                foreach (var o in arr)
                {
                    var child = o as IComponent2;
                    if (child == null) continue;
                    var mdl = child.GetModelDoc2() as IModelDoc2;
                    if (mdl != null && (mdl as IAssemblyDoc) != null)
                    {
                        AddAssemblyToList(child, list, seen);
                        TraverseAssemblies(child, list, seen);
                    }
                }
            }
        }

        private bool ResolveComponent(IComponent2 comp)
        {
            try
            {
                if (comp == null) return false;
                int state = comp.GetSuppression2();
                // Only attempt to resolve if lightweight
                if (state == (int)swComponentSuppressionState_e.swComponentLightweight ||
                    state == (int)swComponentSuppressionState_e.swComponentFullyLightweight)
                {
                    int nRet = comp.SetSuppression2((int)swComponentSuppressionState_e.swComponentResolved);
                    // 0 = completed, 1 = error, 2 = not performed
                    return nRet == 0;
                }
                // Already resolved or suppressed state not applicable
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ResolveComponent error: " + ex.Message);
                return false;
            }
        }

        private void ProcessAllAssemblies(List<IComponent2> assemblies)
        {
            foreach (var comp in assemblies)
            {
                var mdl = comp.GetModelDoc2() as IModelDoc2;
                var asm = mdl as IAssemblyDoc;
                if (asm != null)
                {
                    try
                    {
                        System.Diagnostics.Debug.WriteLine("Reordering subassembly doc: " + Safe(mdl.GetTitle()));
                        ReorderAssembliesThenParts(asm);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("Reorder error: " + ex.Message);
                    }
                }
            }
        }

        private void ReorderAssembliesThenParts(IAssemblyDoc asm)
        {
            if (asm == null) return;
            var model = (IModelDoc2)asm;
            var cfg = model?.ConfigurationManager?.ActiveConfiguration;
            var root = cfg?.GetRootComponent3(true);
            if (root == null)
            {
                System.Windows.Forms.MessageBox.Show("No root component found.");
                return;
            }

            var childrenObj = root.GetChildren();
            if (!(childrenObj is Array arr) || arr.Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("No child components found under the root component.");
                return;
            }

            var assemblies = new List<IComponent2>();
            var parts = new List<IComponent2>();

            foreach (var o in arr)
            {
                var child = o as IComponent2;
                if (child == null) continue;

                // Try to fix the component in space (select then fix)
                try
                {
                    model.ClearSelection2(true);
                    bool sel = model.Extension.SelectByID2(child.Name2, "COMPONENT", 0, 0, 0, false, 0, null, 0);
                    if (sel)
                    {
                        asm.FixComponent();
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("Failed to select component for fix: " + child.Name2);
                    }
                }
                catch { }

                var childDoc = child.GetModelDoc2() as IModelDoc2;
                if (childDoc != null && (childDoc as IAssemblyDoc) != null)
                {
                    assemblies.Add(child);
                    System.Diagnostics.Debug.WriteLine("ASM: " + Safe(childDoc.GetTitle()));
                }
                else
                {
                    parts.Add(child);
                    if (childDoc != null)
                        System.Diagnostics.Debug.WriteLine("Part: " + Safe(childDoc.GetTitle()));
                }
            }

            // Sort by Name2 (case-insensitive)
            Comparison<IComponent2> cmp = (a, b) => string.Compare(Safe(a?.Name2), Safe(b?.Name2), StringComparison.OrdinalIgnoreCase);
            assemblies.Sort(cmp);
            parts.Sort(cmp);

            // Merge assemblies first, then parts
            var merged = new List<IComponent2>(assemblies.Count + parts.Count);
            merged.AddRange(assemblies);
            merged.AddRange(parts);
            if (merged.Count == 0) return;

            var target = merged[0];
            bool ok = false;
            try
            {
                ok = asm.ReorderComponents(merged.Cast<object>().ToArray(), target, REORDER_AFTER);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ReorderComponents exception: " + ex.Message);
                ok = false;
            }

            if (ok)
            {
                GroupDuplicatesInMergedArray(root, merged, asm, model);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Reordering failed.");
            }
        }

        private void GroupDuplicatesInMergedArray(IComponent2 parentComp, IList<IComponent2> merged, IAssemblyDoc asm, IModelDoc2 model)
        {
            if (parentComp == null || merged == null || merged.Count == 0) return;

            var dict = new Dictionary<string, List<IComponent2>>(StringComparer.OrdinalIgnoreCase);
            foreach (var comp in merged)
            {
                if (comp == null) continue;
                var mdl = comp.GetModelDoc2() as IModelDoc2;
                string key;
                if (mdl != null)
                {
                    string title = Safe(mdl.GetTitle());
                    string cfg = Safe(mdl.ConfigurationManager?.ActiveConfiguration?.Name);
                    key = title + "_" + cfg;
                }
                else
                {
                    key = Safe(comp.Name2);
                }

                if (!dict.TryGetValue(key, out var list))
                {
                    list = new List<IComponent2>();
                    dict[key] = list;
                }
                list.Add(comp);
            }

            var featMgr = model.FeatureManager;
            foreach (var kvp in dict)
            {
                var key = kvp.Key;
                var list = kvp.Value;
                if (list.Count < DUPLICATE_GROUP_THRESHOLD) continue;

                // Ensure we have a folder feature for this group (create if missing)
                IFeature folder = null;
                try { folder = asm.FeatureByName(key) as IFeature; } catch { folder = null; }
                if (folder == null)
                {
                    try { folder = featMgr.InsertFeatureTreeFolder2(1 /* swFeatureTreeFolder_EmptyBefore */); } catch { folder = null; }
                    if (folder != null)
                    {
                        try { folder.Name = key; } catch { }
                        try { model.ForceRebuild3(true); } catch { }
                        try { folder = asm.FeatureByName(key) as IFeature; } catch { }
                    }
                }

                if (folder == null) continue;

                // Select all components in the group
                try
                {
                    model.ClearSelection2(true);
                    foreach (var c in list)
                    {
                        try { c?.Select4(true, null, false); } catch { }
                    }
                }
                catch { }

                // Reorder into the folder
                try
                {
                    bool ok = asm.ReorderComponents(list.Cast<object>().ToArray(), folder, REORDER_LAST_IN_FOLDER);
                    System.Diagnostics.Debug.WriteLine(ok ? ($"Grouped {list.Count} into folder: {key}") : ($"FAILED group: {key}"));
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Reorder into folder exception: " + ex.Message);
                }
            }
        }

        private static string Safe(string s) => s ?? string.Empty;
    }
}
