using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace swcsharpaddin
{
    /// <summary>
    /// Extracts embedded PNG icon resources to temp files for SolidWorks 2022+ icon API.
    /// The modern ICommandGroup.IconList / MainIconList properties accept string arrays
    /// of PNG file paths at multiple resolutions (20, 32, 40, 64, 96, 128).
    /// </summary>
    internal sealed class IconResourceHelper : IDisposable
    {
        private readonly List<string> _tempFiles = new List<string>();

        // SolidWorks 2022+ expects icons at these pixel sizes
        private static readonly int[] IconSizes = { 20, 32, 40, 64, 96, 128 };

        /// <summary>
        /// Extracts toolbar strip PNGs to temp files and returns paths for ICommandGroup.IconList.
        /// Each strip PNG contains N icons side-by-side at the given resolution.
        /// </summary>
        public string[] GetToolbarIconList(Assembly assembly)
        {
            var paths = new string[IconSizes.Length];
            for (int i = 0; i < IconSizes.Length; i++)
            {
                string resourceName = "swcsharpaddin.Icons.ToolbarStrip_" + IconSizes[i] + ".png";
                paths[i] = ExtractResourceToTempFile(assembly, resourceName,
                    "NMAutoPilot_Toolbar_" + IconSizes[i] + ".png");
            }
            return paths;
        }

        /// <summary>
        /// Extracts main/group icon PNGs to temp files and returns paths for ICommandGroup.MainIconList.
        /// Each PNG is a single icon at the given resolution.
        /// </summary>
        public string[] GetMainIconList(Assembly assembly)
        {
            var paths = new string[IconSizes.Length];
            for (int i = 0; i < IconSizes.Length; i++)
            {
                string resourceName = "swcsharpaddin.Icons.MainIcon_" + IconSizes[i] + ".png";
                paths[i] = ExtractResourceToTempFile(assembly, resourceName,
                    "NMAutoPilot_Main_" + IconSizes[i] + ".png");
            }
            return paths;
        }

        private string ExtractResourceToTempFile(Assembly assembly, string resourceName, string fileName)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), fileName);

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                    throw new InvalidOperationException("Icon resource not found: " + resourceName);

                using (var fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                {
                    stream.CopyTo(fileStream);
                }
            }

            _tempFiles.Add(tempPath);
            return tempPath;
        }

        public void Dispose()
        {
            foreach (var path in _tempFiles)
            {
                try
                {
                    if (File.Exists(path))
                        File.Delete(path);
                }
                catch
                {
                    // Best-effort cleanup - temp files will be cleaned by OS eventually
                }
            }
            _tempFiles.Clear();
        }
    }
}
