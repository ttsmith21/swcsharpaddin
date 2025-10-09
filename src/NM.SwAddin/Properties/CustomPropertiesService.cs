using System;
using System.Collections.Generic;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Properties
{
    /// <summary>
    /// Reads all custom properties from a model into the ModelInfo cache, tracks changes, and writes back in batch.
    /// - ReadIntoCache overlays config-scope values on top of global values and marks cache clean
    /// - WritePending writes only Added/Modified/Deleted properties to the chosen scopes and marks cache clean
    /// </summary>
    public sealed class CustomPropertiesService
    {
        public bool ReadIntoCache(IModelDoc2 model, ModelInfo info, bool includeGlobal = true, bool includeConfig = true)
        {
            const string proc = nameof(ReadIntoCache);
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (model == null || info == null)
                {
                    ErrorHandler.HandleError(proc, "Invalid inputs");
                    return false;
                }

                var cache = info.CustomProperties;
                if (cache == null)
                {
                    ErrorHandler.HandleError(proc, "ModelInfo.CustomProperties is null");
                    return false;
                }

                // 1) Global properties
                if (includeGlobal)
                {
                    if (SolidWorksApiWrapper.GetCustomProperties(model, "", out var names, out var types, out var values))
                    {
                        OverlayIntoCache(cache, names, types, values);
                    }
                }

                // 2) Configuration properties (overlay)
                string cfg = info.ConfigurationName ?? string.Empty;
                if (includeConfig && !string.IsNullOrWhiteSpace(cfg))
                {
                    if (SolidWorksApiWrapper.GetCustomProperties(model, cfg, out var names, out var types, out var values))
                    {
                        OverlayIntoCache(cache, names, types, values);
                    }
                }

                // After loading from SW, mark cache as clean baseline
                cache.ConfigurationName = cfg;
                cache.MarkClean();
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                return false;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        /// <summary>
        /// Legacy rule: If configuration is Default (or empty), write to Global only; otherwise write to that configuration only.
        /// </summary>
        public bool WritePending(IModelDoc2 model, ModelInfo info)
        {
            string cfgName = info?.ConfigurationName ?? string.Empty;
            bool isDefault = string.IsNullOrWhiteSpace(cfgName) || string.Equals(cfgName, "Default", StringComparison.OrdinalIgnoreCase);
            bool writeGlobal = isDefault;
            bool writeConfig = !isDefault;
            return WritePending(model, info, writeGlobal, writeConfig);
        }

        public bool WritePending(IModelDoc2 model, ModelInfo info, bool writeGlobal = true, bool writeConfig = false)
        {
            const string proc = nameof(WritePending);
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (model == null || info == null)
                {
                    ErrorHandler.HandleError(proc, "Invalid inputs");
                    return false;
                }

                var cache = info.CustomProperties;
                if (cache == null || !cache.IsDirty)
                {
                    // Nothing to write
                    return true;
                }

                string cfgName = info.ConfigurationName ?? string.Empty;
                var states = cache.GetPropertyStates();
                var types = cache.GetPropertyTypes();
                var values = cache.GetProperties();

                foreach (var kv in states)
                {
                    string name = kv.Key;
                    var state = kv.Value;
                    values.TryGetValue(name, out var valueObj);
                    string value = valueObj?.ToString() ?? string.Empty;
                    var type = types.TryGetValue(name, out var t) ? t : CustomPropertyType.Text;
                    var swType = ToSwType(type);

                    switch (state)
                    {
                        case PropertyState.Unchanged:
                            break;
                        case PropertyState.Added:
                        case PropertyState.Modified:
                            if (writeGlobal)
                                SolidWorksApiWrapper.AddCustomProperty(model, name, (swCustomInfoType_e)swType, value, "");
                            if (writeConfig && !string.IsNullOrWhiteSpace(cfgName))
                                SolidWorksApiWrapper.AddCustomProperty(model, name, (swCustomInfoType_e)swType, value, cfgName);
                            break;
                        case PropertyState.Deleted:
                            if (writeGlobal)
                                SolidWorksApiWrapper.DeleteCustomProperty(model, name, "");
                            if (writeConfig && !string.IsNullOrWhiteSpace(cfgName))
                                SolidWorksApiWrapper.DeleteCustomProperty(model, name, cfgName);
                            break;
                    }
                }

                // Mark caches clean after successful write
                cache.MarkClean();
                info.MarkModelClean();
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                return false;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        private void OverlayIntoCache(CustomPropertyData cache, string[] names, int[] types, string[] values)
        {
            if (names == null || values == null) return;
            int len = Math.Min(names.Length, values.Length);
            for (int i = 0; i < len; i++)
            {
                string n = names[i] ?? string.Empty;
                if (string.IsNullOrWhiteSpace(n)) continue;
                string v = values[i] ?? string.Empty;
                var type = FromSwType(i < (types?.Length ?? 0) ? types[i] : (int)swCustomInfoType_e.swCustomInfoText);
                cache.SetPropertyValue(n, v, type);
            }
        }

        private static int ToSwType(CustomPropertyType t)
        {
            switch (t)
            {
                case CustomPropertyType.Number: return (int)swCustomInfoType_e.swCustomInfoNumber;
                case CustomPropertyType.Date: return (int)swCustomInfoType_e.swCustomInfoDate;
                default: return (int)swCustomInfoType_e.swCustomInfoText;
            }
        }

        private static CustomPropertyType FromSwType(int swType)
        {
            if (swType == (int)swCustomInfoType_e.swCustomInfoNumber) return CustomPropertyType.Number;
            if (swType == (int)swCustomInfoType_e.swCustomInfoDate) return CustomPropertyType.Date;
            return CustomPropertyType.Text;
        }
    }
}
