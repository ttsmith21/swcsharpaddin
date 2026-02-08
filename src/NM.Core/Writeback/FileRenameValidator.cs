using System;
using System.IO;
using System.Text.RegularExpressions;
using NM.Core.Reconciliation.Models;
using NM.Core.Writeback.Models;

namespace NM.Core.Writeback
{
    /// <summary>
    /// Validates file rename operations and generates safe filenames from part numbers.
    /// Pure logic — no SolidWorks dependency. The actual rename (SaveAs/ReplaceReferencedDocument)
    /// is handled by the caller in NM.SwAddin.
    /// </summary>
    public sealed class FileRenameValidator
    {
        private static readonly char[] InvalidChars = Path.GetInvalidFileNameChars();
        private static readonly Regex MultipleSpaces = new Regex(@"\s{2,}", RegexOptions.Compiled);

        /// <summary>
        /// Validates a rename suggestion and returns a result indicating whether it's safe to proceed.
        /// </summary>
        public FileRenameValidation Validate(RenameSuggestion rename)
        {
            if (rename == null)
                return FileRenameValidation.Error("Rename suggestion is null");

            if (string.IsNullOrWhiteSpace(rename.OldPath))
                return FileRenameValidation.Error("Current file path is empty");

            if (string.IsNullOrWhiteSpace(rename.NewPath))
                return FileRenameValidation.Error("New file path is empty");

            // Check old file exists
            if (!File.Exists(rename.OldPath))
                return FileRenameValidation.Error($"Source file not found: {rename.OldPath}");

            // Check new filename is valid
            string newFileName = Path.GetFileName(rename.NewPath);
            if (!IsValidFileName(newFileName))
                return FileRenameValidation.Error($"Invalid filename: {newFileName}");

            // Check destination doesn't already exist
            if (File.Exists(rename.NewPath))
            {
                if (string.Equals(rename.OldPath, rename.NewPath, StringComparison.OrdinalIgnoreCase))
                    return FileRenameValidation.Skip("Source and destination are the same file");

                return FileRenameValidation.Error($"Destination already exists: {rename.NewPath}");
            }

            // Check destination directory exists
            string destDir = Path.GetDirectoryName(rename.NewPath);
            if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                return FileRenameValidation.Error($"Destination directory not found: {destDir}");

            // Warn about affected assemblies
            bool hasAssemblyRefs = rename.AffectedAssemblies != null && rename.AffectedAssemblies.Count > 0;

            return new FileRenameValidation
            {
                IsValid = true,
                OldPath = rename.OldPath,
                NewPath = rename.NewPath,
                HasAssemblyReferences = hasAssemblyRefs,
                AffectedAssemblyCount = hasAssemblyRefs ? rename.AffectedAssemblies.Count : 0,
                Warning = hasAssemblyRefs
                    ? $"{rename.AffectedAssemblies.Count} assembly reference(s) will need updating"
                    : null
            };
        }

        /// <summary>
        /// Generates a safe filename from a part number, preserving the original file extension.
        /// </summary>
        /// <param name="partNumber">The part number to use as the filename.</param>
        /// <param name="originalPath">The original file path (for extension and directory).</param>
        /// <returns>Full path with sanitized filename, or null if part number is invalid.</returns>
        public string GenerateNewPath(string partNumber, string originalPath)
        {
            if (string.IsNullOrWhiteSpace(partNumber) || string.IsNullOrWhiteSpace(originalPath))
                return null;

            string sanitized = SanitizeFileName(partNumber);
            if (string.IsNullOrWhiteSpace(sanitized))
                return null;

            string directory = Path.GetDirectoryName(originalPath);
            string extension = Path.GetExtension(originalPath);

            return Path.Combine(directory ?? "", sanitized + extension);
        }

        /// <summary>
        /// Generates a companion drawing path from a model rename.
        /// E.g., if model renames from A.SLDPRT to B.SLDPRT,
        /// and drawing A.SLDDRW exists, returns B.SLDDRW.
        /// </summary>
        public string GenerateDrawingPath(string newModelPath)
        {
            if (string.IsNullOrWhiteSpace(newModelPath))
                return null;

            string dir = Path.GetDirectoryName(newModelPath);
            string nameWithoutExt = Path.GetFileNameWithoutExtension(newModelPath);

            return Path.Combine(dir ?? "", nameWithoutExt + ".SLDDRW");
        }

        /// <summary>
        /// Sanitizes a string for use as a filename, removing invalid characters
        /// and normalizing whitespace.
        /// </summary>
        public static string SanitizeFileName(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return null;

            string result = input.Trim();

            // Remove invalid filename characters
            foreach (char c in InvalidChars)
            {
                result = result.Replace(c.ToString(), "");
            }

            // Collapse multiple spaces
            result = MultipleSpaces.Replace(result, " ");

            // Remove leading/trailing dots (Windows reserved)
            result = result.Trim('.', ' ');

            return string.IsNullOrWhiteSpace(result) ? null : result;
        }

        /// <summary>
        /// Checks if a filename (without path) contains only valid characters.
        /// </summary>
        public static bool IsValidFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return false;

            if (fileName.IndexOfAny(InvalidChars) >= 0)
                return false;

            // Check for Windows reserved names
            string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName).ToUpperInvariant();
            switch (nameWithoutExt)
            {
                case "CON": case "PRN": case "AUX": case "NUL":
                case "COM1": case "COM2": case "COM3": case "COM4":
                case "COM5": case "COM6": case "COM7": case "COM8": case "COM9":
                case "LPT1": case "LPT2": case "LPT3": case "LPT4":
                case "LPT5": case "LPT6": case "LPT7": case "LPT8": case "LPT9":
                    return false;
            }

            return true;
        }
    }
}
