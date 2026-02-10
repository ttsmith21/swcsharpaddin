using System;
using System.Collections.Generic;
using System.Linq;
using NM.Core;
using NM.Core.ProblemParts;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.UI
{
    /// <summary>
    /// Toggles assembly component colors to highlight problem parts in red.
    /// Saves original appearances so they can be restored on toggle-off.
    /// </summary>
    public sealed class ProblemPartColorizer
    {
        private bool _isActive;
        private readonly Dictionary<string, double[]> _savedAppearances = new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Whether problem coloring is currently applied.</summary>
        public bool IsActive => _isActive;

        /// <summary>
        /// Toggles problem part highlighting on/off.
        /// Returns the number of components affected.
        /// </summary>
        public int Toggle(ISldWorks swApp)
        {
            if (swApp == null) return 0;

            var doc = swApp.ActiveDoc as IModelDoc2;
            if (doc == null || doc.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
                return 0;

            int affected;
            if (_isActive)
            {
                affected = RestoreColors(doc);
                _isActive = false;
            }
            else
            {
                affected = ApplyColors(doc);
                _isActive = affected > 0 || _isActive;
            }

            // Force graphics refresh
            doc.GraphicsRedraw2();

            return affected;
        }

        /// <summary>
        /// Resets state without restoring colors (use when assembly changes).
        /// </summary>
        public void Reset()
        {
            _savedAppearances.Clear();
            _isActive = false;
        }

        private int ApplyColors(IModelDoc2 doc)
        {
            var problems = ProblemPartManager.Instance.GetProblemParts();
            if (problems.Count == 0) return 0;

            // Build lookup set of problem file paths
            var problemPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in problems)
            {
                if (!string.IsNullOrEmpty(p.FilePath))
                    problemPaths.Add(p.FilePath);
            }

            var components = GetAllComponents(doc);
            int count = 0;

            foreach (var comp in components)
            {
                try
                {
                    string compPath = comp.GetPathName();
                    if (string.IsNullOrEmpty(compPath)) continue;
                    if (!problemPaths.Contains(compPath)) continue;

                    string compName = comp.Name2 ?? compPath;

                    // Save current appearance before overwriting
                    if (!_savedAppearances.ContainsKey(compName))
                    {
                        var current = (double[])comp.GetMaterialPropertyValues2(
                            (int)swInConfigurationOpts_e.swThisConfiguration, null);
                        if (current != null)
                            _savedAppearances[compName] = (double[])current.Clone();
                        else
                            _savedAppearances[compName] = null;
                    }

                    // Apply red: [R, G, B, Ambient, Diffuse, Specular, Shininess, Transparency, Emission]
                    var red = new double[] { 1.0, 0.0, 0.0, 0.5, 1.0, 0.5, 0.3, 0.0, 0.0 };
                    comp.SetMaterialPropertyValues2(red,
                        (int)swInConfigurationOpts_e.swThisConfiguration, null);

                    count++;
                    ErrorHandler.DebugLog($"[Colorizer] Applied red to: {compName}");
                }
                catch (Exception ex)
                {
                    ErrorHandler.DebugLog($"[Colorizer] Error coloring component: {ex.Message}");
                }
            }

            ErrorHandler.DebugLog($"[Colorizer] Applied red to {count} components");
            return count;
        }

        private int RestoreColors(IModelDoc2 doc)
        {
            var components = GetAllComponents(doc);
            int count = 0;

            foreach (var comp in components)
            {
                try
                {
                    string compName = comp.Name2;
                    if (string.IsNullOrEmpty(compName)) continue;

                    if (_savedAppearances.TryGetValue(compName, out double[] saved) && saved != null)
                    {
                        comp.SetMaterialPropertyValues2(saved,
                            (int)swInConfigurationOpts_e.swThisConfiguration, null);
                        count++;
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandler.DebugLog($"[Colorizer] Error restoring component: {ex.Message}");
                }
            }

            _savedAppearances.Clear();
            ErrorHandler.DebugLog($"[Colorizer] Restored {count} components");
            return count;
        }

        private static List<IComponent2> GetAllComponents(IModelDoc2 doc)
        {
            var result = new List<IComponent2>();
            try
            {
                var config = doc.ConfigurationManager?.ActiveConfiguration;
                if (config == null) return result;

                var root = (IComponent2)config.GetRootComponent3(true);
                if (root == null) return result;

                CollectComponents(root, result);
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[Colorizer] Error getting components: {ex.Message}");
            }
            return result;
        }

        private static void CollectComponents(IComponent2 parent, List<IComponent2> result)
        {
            var childrenRaw = (object[])parent.GetChildren();
            if (childrenRaw == null) return;

            foreach (IComponent2 child in childrenRaw)
            {
                // Skip suppressed components
                if (child.GetSuppression2() == (int)swComponentSuppressionState_e.swComponentSuppressed)
                    continue;

                result.Add(child);
                CollectComponents(child, result);
            }
        }
    }
}
