using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using NM.Core.Models;

namespace NM.SwAddin.AssemblyProcessing
{
    public sealed class ComponentCollector
    {
        private readonly Dictionary<string, SwModelInfo> _unique = new Dictionary<string, SwModelInfo>(System.StringComparer.OrdinalIgnoreCase);
        private readonly List<SwModelInfo> _problems = new List<SwModelInfo>();
        private readonly ComponentValidator _validator = new ComponentValidator();

        public sealed class ComponentValidationResult
        {
            public List<SwModelInfo> ValidComponents { get; set; } = new List<SwModelInfo>();
            public List<SwModelInfo> ProblemComponents { get; set; } = new List<SwModelInfo>();
            public Dictionary<string, int> SkippedReasons { get; set; } = new Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);
        }

        public ComponentValidationResult CollectUniqueComponents(IAssemblyDoc assembly)
        {
            _unique.Clear();
            _problems.Clear();
            var res = new ComponentValidationResult();
            if (assembly == null) return res;

            object compsObj = assembly.GetComponents(false);
            var comps = compsObj as object[];
            if (comps == null) return res;

            foreach (var o in comps)
            {
                var comp = o as IComponent2; if (comp == null) continue;
                string filePath = Safe(() => comp.GetPathName()) ?? string.Empty;
                string config = Safe(() => comp.ReferencedConfiguration) ?? string.Empty;
                string uniqueKey = ($"{filePath.ToLowerInvariant()}::{config}");
                if (_unique.ContainsKey(uniqueKey)) continue;

                var modelInfo = new SwModelInfo(filePath, config)
                {
                    ComponentName = Safe(() => comp.Name2) ?? string.Empty
                };

                var v = _validator.ValidateComponent(comp, modelInfo);
                if (v.IsValid)
                {
                    _unique[uniqueKey] = modelInfo;
                    res.ValidComponents.Add(modelInfo);
                }
                else
                {
                    modelInfo.ProblemDescription = v.Reason;
                    res.ProblemComponents.Add(modelInfo);
                    if (!res.SkippedReasons.ContainsKey(v.Reason)) res.SkippedReasons[v.Reason] = 0;
                    res.SkippedReasons[v.Reason]++;
                }
            }
            return res;
        }

        private static T Safe<T>(System.Func<T> f) { try { return f(); } catch { return default(T); } }
    }
}
