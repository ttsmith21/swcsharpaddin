using System;
using System.IO;
using System.Linq;
using System.Reflection;
using SolidWorks.Interop.sldworks;
using NM.Core;
using System.Collections.Generic;
using Microsoft.Win32;

namespace NM.SwAddin
{
    /// <summary>
    /// Adapter to invoke the legacy ExternalStart automation to populate tube-related properties.
    /// Strategy: Load by path (best-effort) -> GetAddInObject by GUID/Title/ProgID -> fallback TypeFromProgID ->
    /// try preferred method names -> brute-force invoke public methods. Emits verbose diagnostics.
    /// </summary>
    public static class ExternalStartAdapter
    {
        // From registry screenshot (HKLM\Software\SolidWorks\AddIns\{GUID})
        private const string AddInGuidWithBraces = "{c8a18d5b-73a8-4007-a489-c9670fc41443}";
        private const string AddInGuidNoBraces = "c8a18d5b-73a8-4007-a489-c9670fc41443";
        private const string AddInTitle = "Extract Profile Data";

        // Legacy macro server ProgID (present per logs)
        private const string FallbackProgId = "ExternalStart.ExternalStart";

        // Candidates passed to ISldWorks.GetAddInObject (order matters)
        private static readonly string[] ProgIdCandidates = new[]
        {
            AddInGuidWithBraces,
            AddInGuidNoBraces,
            AddInTitle,
            // Prior guesses retained as fallback
            "SWExtractDataAddin.SWExtractDataAddin",
            "SWExtractDataAddin.Addin",
            "SWExtractDataAddin.Connect",
            "ExtractDataAddin.ExtractData",
            "ExtractData.SWExtractDataAddin"
        };

        // Preferred method names to try first
        private static readonly string[] PreferredMethodNames =
            { "ExternalStart", "RunExternalStart", "Start", "Run", "Extract", "ExtractProfile", "ExtractData" };

        // Safe keywords for brute-force discovery
        private static readonly string[] SafeKeywords = { "external", "extract", "profile", "start", "run", "tube" };
        private static readonly string[] UnsafeKeywords = { "connect", "disconnect", "attach", "detach", "event", "notify", "handler" };

        public static bool TryRunExternalStart(IModelDoc2 doc) => TryRunExternalStartInternal(null, doc);
        public static bool TryRunExternalStart(ISldWorks swApp, IModelDoc2 doc) => TryRunExternalStartInternal(swApp, doc);

        private static bool TryRunExternalStartInternal(ISldWorks swApp, IModelDoc2 doc)
        {
            const string proc = nameof(ExternalStartAdapter) + "." + nameof(TryRunExternalStart);
            if (doc == null) return false;
            try
            {
                string title = string.Empty; try { title = doc.GetTitle() ?? string.Empty; } catch { }
                ErrorHandler.DebugLog($"[ExternalStart] Begin for '{title}'");

                // Try loading by explicit path (may not be available in all interops)
                TryLoadByPath(swApp);

                // Try registry-guided discovery first (uses Title match)
                if (swApp != null)
                {
                    foreach (var key in DiscoverAddInKeysByTitle(AddInTitle))
                    {
                        try
                        {
                            ErrorHandler.DebugLog($"[ExternalStart] GetAddInObject(registry:'{key}')...");
                            var automation = swApp.GetAddInObject(key);
                            if (automation == null) { ErrorHandler.DebugLog($"[ExternalStart] GetAddInObject('{key}') => null"); continue; }
                            var t = automation.GetType();
                            ErrorHandler.DebugLog($"[ExternalStart] Automation(registry): {t.FullName}");

                            if (t.Name.EndsWith("SwAddin", StringComparison.OrdinalIgnoreCase))
                            {
                                LogPublicMethods(automation);
                                ErrorHandler.DebugLog("[ExternalStart] Skipping invocation on SwAddin shell (no extractor method exposed).");
                                continue;
                            }

                            if (TryInvokePreferred(automation, doc, out var used)) { LogSuccess(key, used); return true; }
                            if (TryBruteInvoke(automation, doc, out used)) { LogSuccess(key, used); return true; }
                        }
                        catch (Exception ex)
                        {
                            ErrorHandler.HandleError(proc, $"Registry discovery invoke error for '{key}'", ex, "Warning");
                        }
                    }
                }

                // 1) Ask SolidWorks for the automation object by GUID/Title/ProgID
                if (swApp != null)
                {
                    foreach (var key in ProgIdCandidates)
                    {
                        try
                        {
                            ErrorHandler.DebugLog($"[ExternalStart] GetAddInObject('{key}')...");
                            var automation = swApp.GetAddInObject(key);
                            if (automation == null) { ErrorHandler.DebugLog($"[ExternalStart] GetAddInObject('{key}') => null"); continue; }

                            var t = automation.GetType();
                            ErrorHandler.DebugLog($"[ExternalStart] Automation: {t.FullName}");

                            // Do not brute-invoke SwAddin shell types; they usually expose lifecycle methods only
                            if (t.Name.EndsWith("SwAddin", StringComparison.OrdinalIgnoreCase))
                            {
                                LogPublicMethods(automation);
                                ErrorHandler.DebugLog("[ExternalStart] Skipping invocation on SwAddin shell (no extractor method exposed).");
                                continue;
                            }

                            if (TryInvokePreferred(automation, doc, out var used)) { LogSuccess(key, used); return true; }
                            if (TryBruteInvoke(automation, doc, out used)) { LogSuccess(key, used); return true; }

                            LogPublicMethods(automation);
                            ErrorHandler.DebugLog("[ExternalStart] No safe method candidates on automation object.");
                        }
                        catch (Exception ex)
                        {
                            ErrorHandler.HandleError(proc, $"GetAddInObject '{key}' error", ex, "Warning");
                        }
                    }
                }
                else
                {
                    ErrorHandler.DebugLog("[ExternalStart] ISldWorks is null; skipping GetAddInObject phase.");
                }

                // 2) Fallback COM class by ProgID
                try
                {
                    // Try discovered ProgIDs from CLSID first
                    foreach (var progId in DiscoverProgIdsByTitle(AddInTitle))
                    {
                        ErrorHandler.DebugLog($"[ExternalStart] TypeFromProgID(discovered '{progId}')...");
                        if (TryCreateAndInvokeByProgId(progId, doc)) return true;
                    }

                    ErrorHandler.DebugLog($"[ExternalStart] TypeFromProgID('{FallbackProgId}')...");
                    if (TryCreateAndInvokeByProgId(FallbackProgId, doc)) return true;
                }
                catch (Exception ex)
                {
                    ErrorHandler.HandleError(proc, "Fallback ProgID invocation failed", ex, "Warning");
                }

                ErrorHandler.DebugLog("[ExternalStart] Completed with no automation available.");
                return false;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "ExternalStart invocation failed", ex, "Warning");
                return false;
            }
        }

        private static bool TryCreateAndInvokeByProgId(string progId, IModelDoc2 doc)
        {
            try
            {
                var t = Type.GetTypeFromProgID(progId, false);
                if (t != null)
                {
                    var instance = Activator.CreateInstance(t);
                    ErrorHandler.DebugLog($"[ExternalStart] Fallback instance: {instance?.GetType().FullName ?? "null"}");
                    if (instance != null)
                    {
                        if (TryInvokePreferred(instance, doc, out var used)) { ErrorHandler.DebugLog($"[ExternalStart] Invoked '{used}' on '{progId}'."); return true; }
                        if (TryBruteInvoke(instance, doc, out used)) { ErrorHandler.DebugLog($"[ExternalStart] Invoked '{used}' on '{progId}' (brute)."); return true; }
                        LogPublicMethods(instance);
                        ErrorHandler.DebugLog("[ExternalStart] No safe method candidates on fallback instance.");
                    }
                }
                else
                {
                    ErrorHandler.DebugLog($"[ExternalStart] ProgID '{progId}' not registered.");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("ExternalStartAdapter.TryCreateAndInvokeByProgId", $"ProgID '{progId}' failed", ex, "Warning");
            }
            return false;
        }

        private static void TryLoadByPath(ISldWorks swApp)
        {
            try
            {
                if (swApp == null) return;
                var path = NM.Core.Configuration.FilePaths.ExtractDataAddInPath;
                if (string.IsNullOrWhiteSpace(path)) { ErrorHandler.DebugLog("[ExternalStart] Add-in path empty."); return; }
                if (!File.Exists(path)) { ErrorHandler.DebugLog($"[ExternalStart] Add-in path not found: '{path}'"); return; }
                var mi = swApp.GetType().GetMethod("LoadAddIn", BindingFlags.Public | BindingFlags.Instance);
                if (mi == null) { ErrorHandler.DebugLog("[ExternalStart] ISldWorks.LoadAddIn not available on this interop."); return; }
                var res = mi.Invoke(swApp, new object[] { path });
                ErrorHandler.DebugLog($"[ExternalStart] LoadAddIn('{path}') => {res ?? "null"}");
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("ExternalStartAdapter.TryLoadByPath", "LoadAddIn failed", ex, "Warning");
            }
        }

        private static bool TryInvokePreferred(object automation, IModelDoc2 doc, out string used)
        {
            used = string.Empty;
            var t = automation?.GetType();
            if (t == null || doc == null) return false;
            foreach (var name in PreferredMethodNames)
            {
                try
                {
                    // try object/IModelDoc2/ModelDoc2 overloads
                    var withDoc = t.GetMethod(name, BindingFlags.Instance | BindingFlags.Public, null, new[] { typeof(object) }, null)
                               ?? t.GetMethod(name, BindingFlags.Instance | BindingFlags.Public, null, new[] { typeof(IModelDoc2) }, null)
                               ?? t.GetMethod(name, BindingFlags.Instance | BindingFlags.Public, null, new[] { typeof(ModelDoc2) }, null);
                    if (withDoc != null)
                    {
                        ErrorHandler.DebugLog($"[ExternalStart] Trying {withDoc.Name}({withDoc.GetParameters()[0].ParameterType.Name})...");
                        withDoc.Invoke(automation, new object[] { doc });
                        used = withDoc.Name;
                        return true;
                    }
                    var noArgs = t.GetMethod(name, BindingFlags.Instance | BindingFlags.Public, null, Type.EmptyTypes, null);
                    if (noArgs != null)
                    {
                        ErrorHandler.DebugLog($"[ExternalStart] Trying {noArgs.Name}()...");
                        noArgs.Invoke(automation, null);
                        used = noArgs.Name;
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandler.HandleError("ExternalStartAdapter.TryInvokePreferred", $"Invoke '{name}' failed", ex, "Warning");
                }
            }
            return false;
        }

        private static bool TryBruteInvoke(object automation, IModelDoc2 doc, out string used)
        {
            used = string.Empty;
            if (automation == null || doc == null) return false;
            try
            {
                var t = automation.GetType();
                var methods = t.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                               .Where(m => !m.IsSpecialName)
                               .ToArray();

                ErrorHandler.DebugLog($"[ExternalStart] Public methods on {t.FullName}: " +
                    string.Join(", ", methods.Select(m => $"{m.Name}({string.Join(",", m.GetParameters().Select(p => p.ParameterType.Name))})")));

                // Filter to safe candidates only
                var candidates = methods.Where(m =>
                {
                    var n = m.Name.ToLowerInvariant();
                    bool hasSafe = SafeKeywords.Any(k => n.Contains(k));
                    bool hasUnsafe = UnsafeKeywords.Any(k => n.Contains(k));
                    return hasSafe && !hasUnsafe;
                }).ToArray();

                if (candidates.Length == 0)
                {
                    ErrorHandler.DebugLog("[ExternalStart] No safe method candidates found on automation object.");
                    return false;
                }

                // Rank likely entrypoints
                var ranked = candidates.OrderByDescending(m =>
                {
                    var n = m.Name.ToLowerInvariant();
                    int s = 0; if (n.Contains("external")) s += 5; if (n.Contains("extract")) s += 5; if (n.Contains("profile")) s += 4; if (n.Contains("start")) s += 3; if (n.Contains("run")) s += 2; if (n.Contains("tube")) s += 2; return s;
                });

                foreach (var m in ranked)
                {
                    var ps = m.GetParameters();
                    try
                    {
                        if (ps.Length == 0)
                        {
                            ErrorHandler.DebugLog($"[ExternalStart] Trying {m.Name}()...");
                            m.Invoke(automation, null);
                            used = m.Name;
                            return true;
                        }
                        if (ps.Length == 1)
                        {
                            ErrorHandler.DebugLog($"[ExternalStart] Trying {m.Name}({ps[0].ParameterType.Name})...");
                            m.Invoke(automation, new object[] { doc });
                            used = m.Name;
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorHandler.HandleError("ExternalStartAdapter.TryBruteInvoke", $"Invoke '{m.Name}' failed", ex, "Warning");
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("ExternalStartAdapter.TryBruteInvoke", "Reflection failed", ex, "Warning");
            }
            return false;
        }

        private static void LogPublicMethods(object instance)
        {
            try
            {
                var t = instance.GetType();
                var methods = t.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                               .Where(m => !m.IsSpecialName)
                               .Select(m => $"{m.Name}({string.Join(",", m.GetParameters().Select(p => p.ParameterType.Name))})");
                ErrorHandler.DebugLog($"[ExternalStart] Methods available on {t.FullName}: {string.Join(", ", methods)}");
            }
            catch { }
        }

        private static void LogSuccess(string key, string used) =>
            ErrorHandler.DebugLog($"[ExternalStart] Invoked '{used}' on '{key}' successfully.");

        // Registry discovery helpers
        private static IEnumerable<string> DiscoverAddInKeysByTitle(string expectedTitle)
        {
            var keys = new List<string>();
            try
            {
                foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                {
                    try
                    {
                        using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                        using (var addins = baseKey.OpenSubKey(@"SOFTWARE\SolidWorks\AddIns", false))
                        {
                            if (addins == null) continue;
                            foreach (var sub in addins.GetSubKeyNames())
                            {
                                using (var k = addins.OpenSubKey(sub, false))
                                {
                                    var title = k?.GetValue("Title") as string;
                                    if (!string.IsNullOrWhiteSpace(title) && title.IndexOf(expectedTitle, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        keys.Add(sub); // sub is the GUID in {..} format
                                        ErrorHandler.DebugLog($"[ExternalStart] Registry match: '{title}' => {sub}");
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return keys.Distinct(StringComparer.OrdinalIgnoreCase);
        }

        private static IEnumerable<string> DiscoverProgIdsByTitle(string expectedTitle)
        {
            var ids = new List<string>();
            try
            {
                foreach (var guid in DiscoverAddInKeysByTitle(expectedTitle))
                {
                    var g = guid.Trim(); g = g.TrimStart('{').TrimEnd('}');
                    foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                    {
                        try
                        {
                            using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, view))
                            using (var clsid = baseKey.OpenSubKey($"CLSID\\{{{g}}}", false))
                            {
                                var prog = clsid?.OpenSubKey("ProgID", false)?.GetValue(null) as string;
                                if (!string.IsNullOrWhiteSpace(prog))
                                {
                                    ids.Add(prog);
                                    ErrorHandler.DebugLog($"[ExternalStart] Discovered ProgID for {guid}: '{prog}'");
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
            return ids.Distinct(StringComparer.OrdinalIgnoreCase);
        }

        private static T Safe<T>(Func<T> f) { try { return f(); } catch { return default(T); } }
    }
