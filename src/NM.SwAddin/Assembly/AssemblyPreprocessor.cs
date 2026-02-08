using System;
using System.Linq;
using System.Text;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using NM.Core.Models;

namespace NM.SwAddin.AssemblyProcessing
{
    public sealed class AssemblyPreprocessor
    {
        private readonly ComponentCollector _collector = new ComponentCollector();

        public sealed class AssemblyProcessingResult
        {
            public int TotalComponents { get; set; }
            public int UniqueComponents { get; set; }
            public int ValidComponents { get; set; }
            public int ProblemComponents { get; set; }
            public System.Collections.Generic.List<SwModelInfo> ComponentsToProcess { get; set; }
            public System.Collections.Generic.List<SwModelInfo> ProblemParts { get; set; }
            public string Summary { get; set; }
        }

        public AssemblyProcessingResult PreprocessAssembly(IModelDoc2 model)
        {
            if (model == null) throw new ArgumentNullException(nameof(model));
            if (model.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                throw new InvalidOperationException("Model is not an assembly");

            var asm = model as IAssemblyDoc;
            var result = new AssemblyProcessingResult
            {
                ComponentsToProcess = new System.Collections.Generic.List<SwModelInfo>(),
                ProblemParts = new System.Collections.Generic.List<SwModelInfo>()
            };

            object[] components = asm.GetComponents(false) as object[];
            result.TotalComponents = components?.Length ?? 0;

            var collect = _collector.CollectUniqueComponents(asm);
            result.UniqueComponents = collect.ValidComponents.Count + collect.ProblemComponents.Count;
            result.ValidComponents = collect.ValidComponents.Count;
            result.ProblemComponents = collect.ProblemComponents.Count;
            result.ComponentsToProcess = collect.ValidComponents;
            result.ProblemParts = collect.ProblemComponents;

            var sb = new StringBuilder();
            sb.AppendLine("Assembly Processing Summary:");
            sb.AppendLine($"  Total Components: {result.TotalComponents}");
            sb.AppendLine($"  Unique Parts: {result.UniqueComponents}");
            sb.AppendLine($"  Valid for Processing: {result.ValidComponents}");
            sb.AppendLine($"  Problem Parts: {result.ProblemComponents}");
            if (collect.SkippedReasons.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Skip Reasons:");
                foreach (var kv in collect.SkippedReasons.OrderByDescending(k => k.Value))
                {
                    sb.AppendLine($"  {kv.Key}: {kv.Value}");
                }
            }
            result.Summary = sb.ToString();
            return result;
        }

        public bool EnsureComponentsResolved(IAssemblyDoc assembly)
        {
            bool allResolved = true;
            var components = assembly?.GetComponents(false) as object[];
            if (components == null) return true;
            foreach (var o in components)
            {
                var comp = o as IComponent2; if (comp == null) continue;
                var state = (swComponentSuppressionState_e)comp.GetSuppression2();
                if (state == swComponentSuppressionState_e.swComponentLightweight || state == swComponentSuppressionState_e.swComponentFullyLightweight)
                {
                    comp.SetSuppression2((int)swComponentSuppressionState_e.swComponentResolved);
                    var newState = (swComponentSuppressionState_e)comp.GetSuppression2();
                    if (newState != swComponentSuppressionState_e.swComponentResolved && newState != swComponentSuppressionState_e.swComponentFullyResolved)
                    {
                        ErrorHandler.HandleError(nameof(AssemblyPreprocessor) + ".EnsureComponentsResolved", $"Could not resolve: {Safe(() => comp.Name2)}", null, ErrorHandler.LogLevel.Warning);
                        allResolved = false;
                    }
                }
            }
            return allResolved;
        }

        /// <summary>
        /// Fixes (grounds) all floating components in an assembly.
        /// Components imported from STEP files are often not fixed/grounded.
        /// </summary>
        public int FixFloatingComponents(IAssemblyDoc assembly)
        {
            var comps = assembly?.GetComponents(false) as object[];
            if (comps == null) return 0;
            int fixedCount = 0;
            foreach (var compObj in comps)
            {
                var comp = compObj as IComponent2;
                if (comp == null || comp.IsSuppressed()) continue;
                if (!comp.IsFixed())
                {
                    comp.Select4(true, null, false);
                    assembly.FixComponent();
                    fixedCount++;
                }
            }
            if (fixedCount > 0)
                ErrorHandler.DebugLog($"[ASSEMBLY] Fixed {fixedCount} floating component(s)");
            return fixedCount;
        }

        private static T Safe<T>(System.Func<T> f) { try { return f(); } catch { return default(T); } }
    }
}
