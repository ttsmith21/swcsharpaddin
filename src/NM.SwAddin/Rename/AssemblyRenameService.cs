using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NM.Core;
using NM.Core.Rename;
using NM.Core.Writeback;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Rename
{
    /// <summary>
    /// Executes approved renames using the SolidWorks RenameDocument API.
    /// Handles reference updates automatically via IRenamedDocumentReferences.
    /// </summary>
    public sealed class AssemblyRenameService
    {
        private readonly ISldWorks _swApp;

        public AssemblyRenameService(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Validates all rename entries before execution.
        /// Returns list of error messages. Empty list means all clear.
        /// </summary>
        public List<string> ValidateRenames(List<RenameEntry> entries)
        {
            var errors = new List<string>();

            // Check for duplicate target names
            var targetNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in entries)
            {
                if (!entry.IsApproved) continue;

                string nameNoExt = entry.FinalName;
                if (string.IsNullOrWhiteSpace(nameNoExt))
                {
                    errors.Add($"Empty name for component: {entry.CurrentFileName}");
                    continue;
                }

                if (!FileRenameValidator.IsValidFileName(nameNoExt + Path.GetExtension(entry.CurrentFileName)))
                {
                    errors.Add($"Invalid filename: {nameNoExt}");
                    continue;
                }

                string key = nameNoExt.ToUpperInvariant();
                if (targetNames.TryGetValue(key, out string existingSource))
                {
                    errors.Add($"Name collision: \"{nameNoExt}\" used by both \"{existingSource}\" and \"{entry.CurrentFileName}\"");
                }
                else
                {
                    targetNames[key] = entry.CurrentFileName;
                }

                // Check if target file already exists on disk (different from current)
                string ext = Path.GetExtension(entry.CurrentFilePath);
                string dir = Path.GetDirectoryName(entry.CurrentFilePath);
                string targetPath = Path.Combine(dir ?? "", nameNoExt + ext);
                if (!string.Equals(entry.CurrentFilePath, targetPath, StringComparison.OrdinalIgnoreCase)
                    && File.Exists(targetPath))
                {
                    errors.Add($"File already exists: {targetPath}");
                }
            }

            return errors;
        }

        /// <summary>
        /// Executes all approved renames on the active assembly.
        /// Uses ModelDocExtension.RenameDocument for in-memory rename with reference tracking.
        /// </summary>
        public RenameWizardResult ExecuteRenames(IModelDoc2 assemblyModel, List<RenameEntry> entries)
        {
            const string proc = nameof(ExecuteRenames);
            ErrorHandler.PushCallStack(proc);
            try
            {
                var result = new RenameWizardResult
                {
                    TotalComponents = entries.Count
                };

                var approved = entries.Where(e => e.IsApproved).ToList();
                if (approved.Count == 0)
                {
                    result.Summary = "No renames approved.";
                    return result;
                }

                var ext = assemblyModel.Extension;

                foreach (var entry in approved)
                {
                    try
                    {
                        string currentNameNoExt = Path.GetFileNameWithoutExtension(entry.CurrentFileName);
                        string newNameNoExt = entry.FinalName;

                        // Skip if name unchanged
                        if (string.Equals(currentNameNoExt, newNameNoExt, StringComparison.OrdinalIgnoreCase))
                        {
                            result.Skipped++;
                            continue;
                        }

                        // Select the component in the assembly so RenameDocument knows what to rename
                        assemblyModel.ClearSelection2(true);
                        bool selected = assemblyModel.Extension.SelectByID2(
                            entry.CurrentFileName, "COMPONENT", 0, 0, 0, false, 0, null, 0);
                        if (!selected)
                        {
                            result.Failed++;
                            result.Errors.Add($"Could not select component: {entry.CurrentFileName}");
                            ErrorHandler.DebugLog($"[Rename] SelectByID2 failed for {entry.CurrentFileName}");
                            continue;
                        }

                        // RenameDocument takes new name (no extension) - component must be selected first
                        // Returns 0 (swRenameDocumentError_None) on success
                        int renameErr = ext.RenameDocument(newNameNoExt);

                        if (renameErr == 0) // swRenameDocumentError_e.swRenameDocumentError_None
                        {
                            result.Renamed++;
                            ErrorHandler.DebugLog($"[Rename] {entry.CurrentFileName} -> {newNameNoExt}");
                        }
                        else
                        {
                            result.Failed++;
                            string errMsg = $"Rename failed for {entry.CurrentFileName}: error code {renameErr}";
                            result.Errors.Add(errMsg);
                            ErrorHandler.DebugLog("[Rename] " + errMsg);
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Failed++;
                        result.Errors.Add($"Exception renaming {entry.CurrentFileName}: {ex.Message}");
                        ErrorHandler.HandleError(proc, $"Error renaming {entry.CurrentFileName}", ex);
                    }
                }

                // Check if SW has pending renamed document references
                try
                {
                    bool hasRenamed = ext.HasRenamedDocuments();
                    if (hasRenamed)
                    {
                        ErrorHandler.DebugLog("[Rename] Renamed document references pending update on save");
                    }
                }
                catch (Exception ex)
                {
                    ErrorHandler.DebugLog($"[Rename] HasRenamedDocuments warning: {ex.Message}");
                }

                // Save the assembly to persist renames
                if (result.Renamed > 0)
                {
                    int saveErr = 0, saveWarn = 0;
                    bool saved = assemblyModel.Save3(
                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                        ref saveErr, ref saveWarn);

                    if (!saved)
                    {
                        result.Errors.Add($"Assembly save warning: errors={saveErr}, warnings={saveWarn}");
                        ErrorHandler.DebugLog($"[Rename] Save assembly: errors={saveErr}, warnings={saveWarn}");
                    }
                }

                result.Summary = $"Renamed {result.Renamed} of {approved.Count} components" +
                    (result.Failed > 0 ? $", {result.Failed} failed" : "") +
                    (result.Skipped > 0 ? $", {result.Skipped} skipped (unchanged)" : "");

                return result;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }
    }
}
