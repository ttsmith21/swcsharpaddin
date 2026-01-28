using System;
using System.IO;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using NM.Core.Models;

namespace NM.SwAddin.AssemblyProcessing
{
    public sealed class ComponentValidator
    {
        public sealed class ValidationResult
        {
            public bool IsValid { get; set; }
            public string Reason { get; set; }
            public swComponentSuppressionState_e State { get; set; }
        }

        public ValidationResult ValidateComponent(IComponent2 component, SwModelInfo modelInfo)
        {
            var result = new ValidationResult { IsValid = true, Reason = string.Empty };
            if (component == null || modelInfo == null)
            {
                return new ValidationResult { IsValid = false, Reason = "Invalid component/model info" };
            }

            try
            {
                // 1) Suppression / lightweight
                var suppression = (swComponentSuppressionState_e)component.GetSuppression2();
                result.State = suppression;
                switch (suppression)
                {
                    case swComponentSuppressionState_e.swComponentSuppressed:
                        result.IsValid = false;
                        result.Reason = "Component is suppressed";
                        return result;

                    case swComponentSuppressionState_e.swComponentLightweight:
                    case swComponentSuppressionState_e.swComponentFullyLightweight:
                        // Try to resolve
                        try { component.SetSuppression2((int)swComponentSuppressionState_e.swComponentResolved); } catch { }
                        var newState = (swComponentSuppressionState_e)component.GetSuppression2();
                        if (newState != swComponentSuppressionState_e.swComponentResolved &&
                            newState != swComponentSuppressionState_e.swComponentFullyResolved)
                        {
                            result.IsValid = false;
                            result.Reason = "Could not resolve lightweight component";
                            return result;
                        }
                        break;
                }

                // 2) Virtual component (contains '^' in path)
                if (!string.IsNullOrWhiteSpace(modelInfo.FilePath) && modelInfo.FilePath.IndexOf('^') >= 0)
                {
                    result.IsValid = false;
                    result.Reason = "Virtual component";
                    return result;
                }

                // 3) Imported (STEP/IGES/others) by extension check
                if (IsImportedComponent(component))
                {
                    result.IsValid = false;
                    result.Reason = "Imported component (STEP/IGES)";
                    return result;
                }

                // 4) File exists
                if (string.IsNullOrWhiteSpace(modelInfo.FilePath) || !File.Exists(modelInfo.FilePath))
                {
                    result.IsValid = false;
                    result.Reason = "File not found";
                    return result;
                }

                // 5) Toolbox
                if (IsToolboxComponent(component))
                {
                    result.IsValid = false;
                    result.Reason = "Toolbox/library component";
                    return result;
                }

                return result; // valid
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(nameof(ComponentValidator) + ".ValidateComponent", "Validation failed", ex, ErrorHandler.LogLevel.Warning);
                return new ValidationResult { IsValid = false, Reason = "Validation error" };
            }
        }

        private static bool IsImportedComponent(IComponent2 component)
        {
            try
            {
                string path = (component?.GetPathName() ?? string.Empty).ToLowerInvariant();
                return path.EndsWith(".step") || path.EndsWith(".stp") || path.EndsWith(".iges") || path.EndsWith(".igs");
            }
            catch { return false; }
        }

        private static bool IsToolboxComponent(IComponent2 component)
        {
            try
            {
                var model = component?.GetModelDoc2() as IModelDoc2;
                if (model == null) return false;
                var mgr = model.Extension?.get_CustomPropertyManager("");
                if (mgr == null) return false;
                string raw, resolved; raw = resolved = string.Empty;
                try { mgr.Get4("IsToolboxPart", false, out raw, out resolved); } catch { }
                var v = (resolved ?? raw ?? string.Empty).Trim();
                return v.Equals("true", System.StringComparison.OrdinalIgnoreCase) || v.Equals("yes", System.StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }
    }
}
