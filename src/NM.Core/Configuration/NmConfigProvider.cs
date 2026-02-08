using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using NM.Core.Config.Tables;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace NM.Core.Config
{
    /// <summary>
    /// Singleton configuration provider. Loads nm-config.json and nm-tables.json at startup,
    /// validates values, and exposes strongly-typed POCOs for the rest of the application.
    ///
    /// Fallback chain:
    ///   1. Explicit configDir path (if supplied)
    ///   2. {addin-dir}/config/ (beside the DLL)
    ///   3. Embedded defaults (the POCO default values compiled into the DLL)
    ///
    /// After Initialize(), all reads are simple property access — zero overhead.
    /// </summary>
    public static class NmConfigProvider
    {
        private static readonly object _lock = new object();
        private static bool _initialized;

        /// <summary>Current business configuration. Never null after Initialize().</summary>
        public static NmConfig Current { get; private set; } = new NmConfig();

        /// <summary>Current lookup tables. Never null after Initialize().</summary>
        public static NmTables Tables { get; private set; } = new NmTables();

        /// <summary>Validation messages from the last load (empty = clean).</summary>
        public static IReadOnlyList<ConfigValidator.ValidationMessage> LastValidation { get; private set; }
            = Array.Empty<ConfigValidator.ValidationMessage>();

        /// <summary>Path that nm-config.json was loaded from, or null if using defaults.</summary>
        public static string LoadedConfigPath { get; private set; }

        /// <summary>Path that nm-tables.json was loaded from, or null if using defaults.</summary>
        public static string LoadedTablesPath { get; private set; }

        /// <summary>
        /// Initialize the configuration system. Call once at add-in startup (ConnectToSW).
        /// Thread-safe; subsequent calls are no-ops.
        /// </summary>
        /// <param name="configDir">
        /// Optional directory containing nm-config.json and nm-tables.json.
        /// If null, searches beside the executing assembly DLL.
        /// </param>
        public static void Initialize(string configDir = null)
        {
            lock (_lock)
            {
                if (_initialized) return;

                var settings = new JsonSerializerSettings
                {
                    MissingMemberHandling = MissingMemberHandling.Ignore,
                    NullValueHandling = NullValueHandling.Ignore,
                    DefaultValueHandling = DefaultValueHandling.Populate
                };

                // Resolve config directory
                string dir = ResolveConfigDir(configDir);

                // Load nm-config.json
                Current = LoadJson<NmConfig>(dir, "nm-config.json", settings, out string configPath);
                LoadedConfigPath = configPath;

                // Load nm-tables.json
                Tables = LoadJson<NmTables>(dir, "nm-tables.json", settings, out string tablesPath);
                LoadedTablesPath = tablesPath;

                // Validate
                LastValidation = ConfigValidator.Validate(Current);

                _initialized = true;
            }
        }

        /// <summary>
        /// Reload configuration from disk. Useful if JSON files are edited while running.
        /// </summary>
        public static void Reload(string configDir = null)
        {
            lock (_lock)
            {
                _initialized = false;
                Initialize(configDir);
            }
        }

        /// <summary>
        /// Reset to compiled defaults. Useful for tests.
        /// </summary>
        public static void ResetToDefaults()
        {
            lock (_lock)
            {
                Current = new NmConfig();
                Tables = new NmTables();
                LoadedConfigPath = null;
                LoadedTablesPath = null;
                LastValidation = Array.Empty<ConfigValidator.ValidationMessage>();
                _initialized = true;
            }
        }

        /// <summary>
        /// Save the current configuration to disk as JSON, then reload.
        /// </summary>
        /// <param name="path">Target file path, or null to use LoadedConfigPath / fallback.</param>
        public static void Save(string path = null)
        {
            lock (_lock)
            {
                string target = path
                    ?? LoadedConfigPath
                    ?? FallbackConfigPath();

                string dir = Path.GetDirectoryName(target);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                var settings = new JsonSerializerSettings
                {
                    ContractResolver = new CamelCasePropertyNamesContractResolver(),
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore
                };

                string json = JsonConvert.SerializeObject(Current, settings);
                File.WriteAllText(target, json);

                LoadedConfigPath = target;
                _initialized = false;
                Initialize(Path.GetDirectoryName(target));
            }
        }

        /// <summary>
        /// Returns the default config file path beside the DLL: {asmDir}/config/nm-config.json.
        /// </summary>
        public static string FallbackConfigPath()
        {
            try
            {
                string asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(asmDir))
                    return Path.Combine(asmDir, "config", "nm-config.json");
            }
            catch { }
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "NM_AutoPilot", "nm-config.json");
        }

        private static string ResolveConfigDir(string explicitDir)
        {
            // 1. Explicit path
            if (!string.IsNullOrWhiteSpace(explicitDir) && Directory.Exists(explicitDir))
                return explicitDir;

            // 2. Beside the executing assembly
            try
            {
                string asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(asmDir))
                {
                    string configDir = Path.Combine(asmDir, "config");
                    if (Directory.Exists(configDir))
                        return configDir;
                }
            }
            catch
            {
                // Assembly location may not be available in some hosting scenarios
            }

            return null;
        }

        private static T LoadJson<T>(string dir, string fileName, JsonSerializerSettings settings, out string loadedPath) where T : new()
        {
            loadedPath = null;

            if (string.IsNullOrEmpty(dir))
                return new T();

            string filePath = Path.Combine(dir, fileName);
            if (!File.Exists(filePath))
                return new T();

            try
            {
                string json = File.ReadAllText(filePath);
                var result = JsonConvert.DeserializeObject<T>(json, settings);
                if (result != null)
                {
                    loadedPath = filePath;
                    return result;
                }
            }
            catch (Exception ex)
            {
                // Log but don't crash — fall back to defaults
                System.Diagnostics.Debug.WriteLine($"[NmConfigProvider] Failed to load {filePath}: {ex.Message}");
            }

            return new T();
        }
    }
}
