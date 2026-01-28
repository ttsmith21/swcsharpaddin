using System;
using System.IO;
using System.Linq;
using Microsoft.Win32;

namespace NM.Core
{
    /// <summary>
    /// Resolves the bend table path to use given processing options and configured defaults.
    /// Falls back gracefully to local paths or K-factor when network paths are unavailable.
    /// Pure Core logic; no SolidWorks types.
    /// </summary>
    public static class BendTableResolver
    {
        /// <summary>
        /// Returns a file path to a bend table, or Configuration.FilePaths.BendTableNone ("-1") to signal K-factor.
        /// </summary>
        public static string Resolve(ProcessingOptions options)
        {
            try
            {
                if (options == null)
                {
                    NM.Core.ErrorHandler.DebugLog("[BendTableResolver] Options is null -> K-Factor");
                    return Configuration.FilePaths.BendTableNone;
                }

                // If explicitly set on options and exists (with variants), honor it
                if (!string.IsNullOrWhiteSpace(options.BendTable) &&
                    !string.Equals(options.BendTable, Configuration.FilePaths.BendTableNone, StringComparison.OrdinalIgnoreCase))
                {
                    NM.Core.ErrorHandler.DebugLog($"[BendTableResolver] Requested: '{options.BendTable}'");
                    var resolved = TryResolveExistingPathVariants(options.BendTable);
                    if (!string.IsNullOrWhiteSpace(resolved))
                    {
                        NM.Core.ErrorHandler.DebugLog($"[BendTableResolver] Using requested(resolved): '{resolved}'");
                        WarnIfXlsx(resolved);
                        return resolved;
                    }
                    NM.Core.ErrorHandler.HandleError("BendTableResolver", $"Requested bend table not found: '{options.BendTable}'", null, ErrorHandler.LogLevel.Warning);
                }

                // Aluminum: always use K-factor
                if (options.MaterialCategory == MaterialCategoryKind.Aluminum)
                {
                    NM.Core.ErrorHandler.DebugLog("[BendTableResolver] Material=Aluminum -> K-Factor");
                    return Configuration.FilePaths.BendTableNone;
                }

                // Candidate files by config
                string[] candidates = Array.Empty<string>();
                if (options.MaterialCategory == MaterialCategoryKind.StainlessSteel)
                {
                    candidates = new[] { Configuration.FilePaths.BendTableSs, Configuration.FilePaths.BendTableSsLocal };
                }
                else if (options.MaterialCategory == MaterialCategoryKind.CarbonSteel)
                {
                    candidates = new[] { Configuration.FilePaths.BendTableCs, Configuration.FilePaths.BendTableCsLocal };
                }

                NM.Core.ErrorHandler.DebugLog($"[BendTableResolver] Candidates: [{string.Join(", ", candidates.Where(c => !string.IsNullOrWhiteSpace(c)))}]");

                foreach (var p in candidates)
                {
                    if (string.IsNullOrWhiteSpace(p)) continue;
                    NM.Core.ErrorHandler.DebugLog($"[BendTableResolver] Check: '{p}'");
                    var resolved = TryResolveExistingPathVariants(p);
                    if (!string.IsNullOrWhiteSpace(resolved)) { WarnIfXlsx(resolved); return resolved; }
                }

                // Prepare keyword filters
                string[] ssKeys = new[] { "stainless", "ss" };
                string[] csKeys = new[] { "carbon", "mild", "steel", "cs" };
                var keys = (options.MaterialCategory == MaterialCategoryKind.StainlessSteel) ? ssKeys : csKeys;

                // Helper: scan a directory for .xls* and choose best match
                Func<string, string> scanDir = (string baseDir) =>
                {
                    try
                    {
                        if (string.IsNullOrWhiteSpace(baseDir))
                        {
                            NM.Core.ErrorHandler.DebugLog("[BendTableResolver] scanDir: baseDir is null/empty");
                            return null;
                        }

                        if (!Directory.Exists(baseDir))
                        {
                            var unc = TryMapDriveToUnc(baseDir);
                            if (!string.IsNullOrWhiteSpace(unc) && Directory.Exists(unc))
                            {
                                NM.Core.ErrorHandler.DebugLog($"[BendTableResolver] scanDir: mapped '{baseDir}' -> '{unc}'");
                                baseDir = unc;
                            }
                        }

                        if (!Directory.Exists(baseDir))
                        {
                            NM.Core.ErrorHandler.DebugLog($"[BendTableResolver] scanDir: directory does not exist: '{baseDir}'");
                            return null;
                        }

                        NM.Core.ErrorHandler.DebugLog($"[BendTableResolver] Scanning dir: '{baseDir}' for .xls/.xlsx");
                        var files = Directory.EnumerateFiles(baseDir, "*.xls*")
                                             .Take(200)
                                             .ToArray();
                        NM.Core.ErrorHandler.DebugLog($"[BendTableResolver] scanDir: found {files.Length} candidate files");
                        if (files.Length == 0) return null;
                        var match = files.FirstOrDefault(f => keys.Any(k => Path.GetFileName(f).IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                                   ?? files.FirstOrDefault();
                        NM.Core.ErrorHandler.DebugLog($"[BendTableResolver] scanDir: chosen='{match}'");
                        WarnIfXlsx(match);
                        return match;
                    }
                    catch (Exception ex)
                    {
                        NM.Core.ErrorHandler.HandleError("BendTableResolver", $"scanDir exception for '{baseDir}'", ex, ErrorHandler.LogLevel.Warning);
                        return null;
                    }
                };

                // Scan configured parent directories (network/local)
                foreach (var p in candidates)
                {
                    try
                    {
                        if (string.IsNullOrWhiteSpace(p)) continue;
                        var d = Path.GetDirectoryName(p);
                        var found = scanDir(d);
                        if (!string.IsNullOrWhiteSpace(found))
                        {
                            NM.Core.ErrorHandler.HandleError("BendTableResolver", $"Using discovered bend table: {found}", null, ErrorHandler.LogLevel.Warning);
                            return found;
                        }
                    }
                    catch (Exception ex)
                    {
                        NM.Core.ErrorHandler.HandleError("BendTableResolver", "Error probing candidate parent dir", ex, ErrorHandler.LogLevel.Warning);
                    }
                }

                // Discover SOLIDWORKS install folders from registry (multiple versions, both hives)
                var installDirs = GetSolidWorksInstallDirs();
                NM.Core.ErrorHandler.DebugLog($"[BendTableResolver] SW install dirs: {installDirs.Length} -> [{string.Join("; ", installDirs)}]");

                // Known subfolders to probe
                string[] subDirs = new[]
                {
                    Path.Combine("lang","english","Sheet Metal Bend Tables"),
                    Path.Combine("lang","english","Sheet Metal Gauge Tables"),
                };

                foreach (var baseDir in installDirs)
                {
                    foreach (var sub in subDirs)
                    {
                        var full = Path.Combine(baseDir, sub);
                        var found = scanDir(full);
                        if (!string.IsNullOrWhiteSpace(found))
                        {
                            NM.Core.ErrorHandler.HandleError("BendTableResolver", $"Using discovered SW bend table: {found}", null, ErrorHandler.LogLevel.Warning);
                            return found;
                        }
                    }
                }

                NM.Core.ErrorHandler.DebugLog("[BendTableResolver] No table found -> K-Factor");
                return Configuration.FilePaths.BendTableNone;
            }
            catch (Exception ex)
            {
                NM.Core.ErrorHandler.HandleError("BendTableResolver", "Resolver exception -> K-Factor", ex, ErrorHandler.LogLevel.Warning);
                return Configuration.FilePaths.BendTableNone;
            }
        }

        private static string[] GetSolidWorksInstallDirs()
        {
            var list = new System.Collections.Generic.List<string>();
            try
            {
                // Typical versions range; adjust as needed
                var versions = new[] { "2018","2019","2020","2021","2022","2023","2024","2025" };
                foreach (var v in versions)
                {
                    TryAddKey(list, $"SOFTWARE\\SolidWorks\\SOLIDWORKS {v}");
                    TryAddKey(list, $"SOFTWARE\\WOW6432Node\\SolidWorks\\SOLIDWORKS {v}");
                }
            }
            catch { }
            return list.Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        }

        private static void TryAddKey(System.Collections.Generic.List<string> list, string subKey)
        {
            try
            {
                using (var k = Registry.LocalMachine.OpenSubKey(subKey))
                {
                    var folder = k?.GetValue("Folder") as string;
                    if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
                    {
                        list.Add(folder);
                    }
                }
            }
            catch { }
        }

        private static string TryResolveExistingPathVariants(string p)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(p)) return null;

                // as-is
                if (File.Exists(p)) return p;

                // try UNC mapping for path
                var unc = TryMapDriveToUnc(p);
                if (!string.IsNullOrWhiteSpace(unc) && File.Exists(unc))
                {
                    NM.Core.ErrorHandler.DebugLog($"[BendTableResolver] Mapped drive -> UNC: '{p}' -> '{unc}'");
                    return unc;
                }

                // try alternate extension between .xls and .xlsx
                var ext = Path.GetExtension(p);
                if (!string.IsNullOrWhiteSpace(ext))
                {
                    string alt = null;
                    if (ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)) alt = Path.ChangeExtension(p, ".xls");
                    else if (ext.Equals(".xls", StringComparison.OrdinalIgnoreCase)) alt = Path.ChangeExtension(p, ".xlsx");
                    if (!string.IsNullOrWhiteSpace(alt))
                    {
                        if (File.Exists(alt)) return alt;
                        var altUnc = TryMapDriveToUnc(alt);
                        if (!string.IsNullOrWhiteSpace(altUnc) && File.Exists(altUnc))
                        {
                            NM.Core.ErrorHandler.DebugLog($"[BendTableResolver] Mapped drive alt-ext -> UNC: '{alt}' -> '{altUnc}'");
                            return altUnc;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                NM.Core.ErrorHandler.HandleError("BendTableResolver", $"TryResolveExistingPathVariants exception for '{p}'", ex, ErrorHandler.LogLevel.Warning);
            }
            return null;
        }

        private static string TryMapDriveToUnc(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path)) return null;
                // Expect pattern 'X:\...'
                if (path.Length >= 3 && path[1] == ':' && (path[2] == '\\' || path[2] == '/'))
                {
                    char drive = char.ToUpperInvariant(path[0]);
                    using (var k = Registry.CurrentUser.OpenSubKey($"Network\\{drive}"))
                    {
                        var remote = k?.GetValue("RemotePath") as string;
                        if (!string.IsNullOrWhiteSpace(remote))
                        {
                            var suffix = path.Substring(3).Replace('/', '\\');
                            var unc = remote.TrimEnd('\\') + "\\" + suffix;
                            return unc;
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        private static void WarnIfXlsx(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path)) return;
                var ext = Path.GetExtension(path);
                if (ext != null && ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    NM.Core.ErrorHandler.HandleError("BendTableResolver", $"Selected bend table is .xlsx: '{path}'. Some SOLIDWORKS versions expect .xls.", null, ErrorHandler.LogLevel.Warning);
                }
            }
            catch { }
        }
    }
}
