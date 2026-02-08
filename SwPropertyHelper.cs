using System;
using NM.Core;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin
{
    /// <summary>
    /// Custom property CRUD operations for SolidWorks models.
    /// Extracted from SolidWorksApiWrapper for single-responsibility.
    /// </summary>
    public static class SwPropertyHelper
    {
        /// <summary>
        /// Adds a custom property to a SolidWorks model.
        /// </summary>
        public static bool AddCustomProperty(IModelDoc2 swModel, string propName, swCustomInfoType_e propType, string propValue, string configName)
        {
            const string procName = "AddCustomProperty";
            ErrorHandler.PushCallStack(procName);
            PerformanceTracker.Instance.StartTimer("CustomProperty_Write");
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return false;
                if (!SolidWorksApiWrapper.ValidateString(propName, procName, "property name")) return false;

                var mgr = swModel.Extension.get_CustomPropertyManager(configName);
                int addResult = mgr.Add3(propName, (int)propType, propValue, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);

                // Verify by querying back since return codes vary by version/options
                bool exists = TryGetProperty(mgr, propName, out var currentValue);
                bool ok = exists; // we only require it to exist; value may be evaluated text

                if (!ok)
                {
                    // Fallback attempts
                    int setResult = mgr.Set2(propName, propValue);
                    exists = TryGetProperty(mgr, propName, out currentValue);
                    if (!exists)
                    {
                        try
                        {
                            int add2Result = mgr.Add2(propName, (int)propType, propValue);
                        }
                        catch { }
                        exists = TryGetProperty(mgr, propName, out currentValue);
                    }

                    ErrorHandler.DebugLog($"AddCustomProperty failed (code={addResult}) for '{propName}'.");
                }
                return ok || exists;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception: {ex.Message}", ex, ErrorHandler.LogLevel.Error, $"Property: {propName}");
                return false;
            }
            finally
            {
                PerformanceTracker.Instance.StopTimer("CustomProperty_Write");
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Sets the value of an existing custom property.
        /// </summary>
        public static bool SetCustomProperty(IModelDoc2 swModel, string propName, string propValue, string configName)
        {
            const string procName = "SetCustomProperty";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return false;
                if (!SolidWorksApiWrapper.ValidateString(propName, procName, "property name")) return false;

                var mgr = swModel.Extension.get_CustomPropertyManager(configName);
                int setResult = mgr.Set2(propName, propValue);

                // Verify
                bool exists = TryGetProperty(mgr, propName, out var currentValue);
                if (!exists)
                {
                    // If Set2 failed due to not existing, try Add2
                    try { mgr.Add2(propName, (int)swCustomInfoType_e.swCustomInfoText, propValue); } catch { }
                    exists = TryGetProperty(mgr, propName, out currentValue);
                }
                return exists;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception: {ex.Message}", ex, ErrorHandler.LogLevel.Error, $"Property: {propName}");
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Retrieves all custom properties from a SolidWorks model.
        /// </summary>
        public static bool GetCustomProperties(IModelDoc2 swModel, string configName, out string[] propNames, out int[] propTypes, out string[] propValues)
        {
            const string procName = "GetCustomProperties";
            propNames = Array.Empty<string>();
            propTypes = Array.Empty<int>();
            propValues = Array.Empty<string>();

            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return false;

                var mgr = swModel.Extension.get_CustomPropertyManager(configName);
                object namesObj = null, typesObj = null, valuesObj = null, resolved = null, links = null;
                mgr.GetAll3(ref namesObj, ref typesObj, ref valuesObj, ref resolved, ref links);

                var namesArr = namesObj as object[] ?? namesObj as string[];
                var typesArr = typesObj as object[];
                var valuesArr = valuesObj as object[] ?? valuesObj as string[];

                if (namesArr == null || typesArr == null || valuesArr == null)
                {
                    return true; // no properties
                }

                int len = Math.Min(namesArr.Length, Math.Min(typesArr.Length, valuesArr.Length));
                propNames = new string[len];
                propTypes = new int[len];
                propValues = new string[len];
                for (int i = 0; i < len; i++)
                {
                    object n = (namesArr is string[]) ? ((string[])namesArr)[i] : namesArr[i];
                    object v = (valuesArr is string[]) ? ((string[])valuesArr)[i] : valuesArr[i];
                    propNames[i] = n?.ToString() ?? string.Empty;
                    propTypes[i] = Convert.ToInt32(typesArr[i] ?? 0);
                    propValues[i] = v?.ToString() ?? string.Empty;
                }
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception: {ex.Message}", ex, ErrorHandler.LogLevel.Error);
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Deletes a custom property from a SolidWorks model.
        /// </summary>
        public static bool DeleteCustomProperty(IModelDoc2 swModel, string propName, string configName)
        {
            const string procName = "DeleteCustomProperty";
            ErrorHandler.PushCallStack(procName);
            try
            {
                if (!SolidWorksApiWrapper.ValidateModel(swModel, procName)) return false;
                if (!SolidWorksApiWrapper.ValidateString(propName, procName, "property name")) return false;

                var mgr = swModel.Extension.get_CustomPropertyManager(configName);
                int delResult = mgr.Delete2(propName);

                bool exists = TryGetProperty(mgr, propName, out var _);
                bool ok = !exists;
                if (!ok)
                {
                    ErrorHandler.DebugLog($"DeleteCustomProperty failed (code={delResult}) for '{propName}'.");
                }
                return ok;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(procName, $"Exception: {ex.Message}", ex, ErrorHandler.LogLevel.Error, $"Property: {propName}");
                return false;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Gets a custom property value from a model.
        /// </summary>
        public static string GetCustomPropertyValue(IModelDoc2 model, string propName, string configName = "")
        {
            if (model == null || string.IsNullOrEmpty(propName)) return string.Empty;
            PerformanceTracker.Instance.StartTimer("CustomProperty_Read");
            try
            {
                var ext = model.Extension;
                if (ext == null) return string.Empty;

                var mgr = string.IsNullOrEmpty(configName)
                    ? ext.get_CustomPropertyManager(string.Empty)
                    : ext.get_CustomPropertyManager(configName);

                if (mgr == null) return string.Empty;

                string val = string.Empty;
                string resolved = string.Empty;
                bool wasResolved = false;
                mgr.Get5(propName, false, out val, out resolved, out wasResolved);
                return resolved ?? val ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                PerformanceTracker.Instance.StopTimer("CustomProperty_Read");
            }
        }

        internal static bool TryGetProperty(ICustomPropertyManager mgr, string propName, out string val)
        {
            val = string.Empty;
            try
            {
                object nObj = null, tObj = null, vObj = null, rObj = null, lObj = null;
                mgr.GetAll3(ref nObj, ref tObj, ref vObj, ref rObj, ref lObj);
                var names = nObj as object[] ?? nObj as string[];
                var values = vObj as object[] ?? vObj as string[];
                if (names == null || values == null) return false;
                for (int i = 0; i < Math.Min(names.Length, values.Length); i++)
                {
                    var n = (names is string[]) ? ((string[])names)[i] : names[i]?.ToString();
                    if (string.Equals(n, propName, StringComparison.OrdinalIgnoreCase))
                    {
                        val = (values is string[]) ? ((string[])values)[i] : values[i]?.ToString();
                        return true;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }
    }
}
