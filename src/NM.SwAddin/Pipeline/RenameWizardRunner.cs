using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NM.Core;
using NM.Core.Pdf;
using NM.Core.Rename;
using NM.SwAddin.Rename;
using NM.SwAddin.UI;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Pipeline
{
    /// <summary>
    /// Orchestrates the Rename Wizard pipeline:
    ///   1. Validate active document is an assembly
    ///   2. Collect unique components
    ///   3. Find/browse for companion PDF
    ///   4. Extract BOM from PDF (text extraction)
    ///   5. Match BOM rows to components (heuristic + optional AI)
    ///   6. Show RenameWizardForm for user review
    ///   7. Validate and execute approved renames
    /// </summary>
    public sealed class RenameWizardRunner
    {
        private readonly ISldWorks _swApp;

        public RenameWizardRunner(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        /// <summary>
        /// Runs the rename wizard on the active assembly.
        /// </summary>
        /// <returns>Result summary, or null if cancelled.</returns>
        public RenameWizardResult RunOnActiveAssembly()
        {
            const string proc = nameof(RunOnActiveAssembly);
            ErrorHandler.PushCallStack(proc);
            try
            {
                // 1. Validate active doc is assembly
                var model = _swApp.ActiveDoc as IModelDoc2;
                if (model == null)
                {
                    MessageBox.Show("No active document. Open an assembly first.",
                        "Rename Wizard", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return null;
                }

                int docType = model.GetType();
                if (docType != (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    MessageBox.Show("Rename Wizard works on assemblies only. Please open an assembly.",
                        "Rename Wizard", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return null;
                }

                string modelPath = model.GetPathName();
                if (string.IsNullOrEmpty(modelPath))
                {
                    MessageBox.Show("Please save the assembly first.",
                        "Rename Wizard", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return null;
                }

                // 2. Collect components
                var assemblyDoc = (IAssemblyDoc)model;
                var componentInfos = CollectComponentInfos(assemblyDoc);

                if (componentInfos.Count == 0)
                {
                    MessageBox.Show("No components found in the assembly.",
                        "Rename Wizard", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return null;
                }

                // 3. Find/browse for companion PDF
                string pdfPath = PdfDrawingAnalyzer.FindCompanionPdf(modelPath);
                if (string.IsNullOrEmpty(pdfPath))
                {
                    var browseResult = MessageBox.Show(
                        $"No companion PDF found for:\n{Path.GetFileName(modelPath)}\n\n" +
                        "Would you like to browse for a PDF?",
                        "Rename Wizard", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (browseResult == DialogResult.Yes)
                        pdfPath = BrowseForPdf(modelPath);

                    if (string.IsNullOrEmpty(pdfPath))
                        return null;
                }

                // 4. Extract BOM from PDF text
                List<BomRow> bomRows;
                try
                {
                    bomRows = ExtractBomFromPdf(pdfPath);
                }
                catch (Exception ex)
                {
                    ErrorHandler.HandleError(proc, "BOM extraction failed", ex);
                    MessageBox.Show($"Failed to extract BOM from PDF:\n{ex.Message}\n\nYou can still rename manually.",
                        "Rename Wizard", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    bomRows = new List<BomRow>();
                }

                // 5. Match BOM rows to components
                var matcher = new BomComponentMatcher();
                var entries = matcher.Match(bomRows, componentInfos);

                int matched = entries.Count(e => e.Confidence > 0);
                ErrorHandler.DebugLog($"[RenameWizard] {matched}/{entries.Count} components matched to BOM rows");

                // 6. Show wizard
                var wizard = new RenameWizardForm(
                    entries,
                    Path.GetFileName(modelPath),
                    Path.GetFileName(pdfPath));

                // Wire up click-to-highlight
                wizard.OnRowSelected = entry => HighlightComponent(model, assemblyDoc, entry);

                if (wizard.ShowDialog() != DialogResult.OK)
                    return null;

                var approved = wizard.ApprovedRenames;
                if (approved == null || approved.Count == 0)
                {
                    MessageBox.Show("No renames selected.",
                        "Rename Wizard", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return null;
                }

                // 7. Validate
                var renameService = new AssemblyRenameService(_swApp);
                var validationErrors = renameService.ValidateRenames(approved);
                if (validationErrors.Count > 0)
                {
                    string errMsg = string.Join("\n", validationErrors.Take(10));
                    if (validationErrors.Count > 10)
                        errMsg += $"\n...and {validationErrors.Count - 10} more";

                    var proceed = MessageBox.Show(
                        $"Validation warnings:\n\n{errMsg}\n\nProceed anyway?",
                        "Rename Wizard", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                    if (proceed != DialogResult.Yes)
                        return null;
                }

                // 8. Execute renames
                var result = renameService.ExecuteRenames(model, approved);

                // 9. Show summary
                string summary = result.Summary;
                if (result.Errors.Count > 0)
                    summary += "\n\nErrors:\n" + string.Join("\n", result.Errors.Take(5));

                MessageBox.Show(summary, "Rename Wizard - Complete",
                    MessageBoxButtons.OK,
                    result.Failed > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);

                return result;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                MessageBox.Show($"Rename wizard failed: {ex.Message}",
                    "Rename Wizard", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        private List<ComponentInfo> CollectComponentInfos(IAssemblyDoc assembly)
        {
            var result = new List<ComponentInfo>();
            var instanceCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // First pass: count instances per file
            object compsObj = assembly.GetComponents(false);
            var comps = compsObj as object[];
            if (comps == null) return result;

            foreach (var o in comps)
            {
                var comp = o as IComponent2;
                if (comp == null) continue;

                string path = SafeGet(() => comp.GetPathName());
                if (string.IsNullOrEmpty(path)) continue;
                if (SafeGet(() => comp.IsSuppressed())) continue;

                string key = path.ToUpperInvariant();
                if (!instanceCounts.ContainsKey(key))
                    instanceCounts[key] = 0;
                instanceCounts[key]++;
            }

            // Second pass: build unique list
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int index = 0;

            foreach (var o in comps)
            {
                var comp = o as IComponent2;
                if (comp == null) continue;

                string path = SafeGet(() => comp.GetPathName());
                if (string.IsNullOrEmpty(path)) continue;
                if (SafeGet(() => comp.IsSuppressed())) continue;

                string config = SafeGet(() => comp.ReferencedConfiguration) ?? "";
                string uniqueKey = $"{path.ToUpperInvariant()}::{config}";

                if (seen.Contains(uniqueKey)) continue;
                seen.Add(uniqueKey);

                string fileName = Path.GetFileName(path);
                string pathKey = path.ToUpperInvariant();

                result.Add(new ComponentInfo
                {
                    Index = index++,
                    FileName = fileName,
                    FilePath = path,
                    Configuration = config,
                    InstanceCount = instanceCounts.TryGetValue(pathKey, out int cnt) ? cnt : 1
                });
            }

            return result;
        }

        private List<BomRow> ExtractBomFromPdf(string pdfPath)
        {
            // Use PdfPig text extraction to find BOM table patterns
            var textExtractor = new PdfTextExtractor();
            var pages = textExtractor.ExtractText(pdfPath);

            var bomRows = new List<BomRow>();
            int itemNum = 1;

            // Look for BOM-like patterns in extracted text
            // Common BOM formats: "ITEM | PART NUMBER | DESCRIPTION | QTY"
            foreach (var page in pages)
            {
                var lines = (page.FullText ?? "").Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                bool inBomSection = false;

                foreach (var line in lines)
                {
                    string trimmed = line.Trim();

                    // Detect BOM header
                    if (IsBomHeader(trimmed))
                    {
                        inBomSection = true;
                        continue;
                    }

                    if (!inBomSection) continue;

                    // Try to parse a BOM row
                    var row = TryParseBomRow(trimmed, itemNum);
                    if (row != null)
                    {
                        bomRows.Add(row);
                        itemNum++;
                    }
                    else if (inBomSection && string.IsNullOrWhiteSpace(trimmed))
                    {
                        // Empty line might signal end of BOM
                        break;
                    }
                }
            }

            ErrorHandler.DebugLog($"[RenameWizard] Extracted {bomRows.Count} BOM rows from PDF text");
            return bomRows;
        }

        private static bool IsBomHeader(string line)
        {
            string upper = line.ToUpperInvariant();
            // Look for common BOM header patterns
            int matchCount = 0;
            if (upper.Contains("ITEM")) matchCount++;
            if (upper.Contains("PART") || upper.Contains("DRAWING")) matchCount++;
            if (upper.Contains("DESCRIPTION") || upper.Contains("NAME")) matchCount++;
            if (upper.Contains("QTY") || upper.Contains("QUANTITY")) matchCount++;
            return matchCount >= 2;
        }

        private static BomRow TryParseBomRow(string line, int fallbackItemNum)
        {
            if (string.IsNullOrWhiteSpace(line)) return null;

            // Split by common delimiters (tab, multiple spaces, pipe)
            var parts = line.Split(new[] { '\t', '|' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
            {
                // Try splitting by 2+ spaces
                parts = System.Text.RegularExpressions.Regex.Split(line.Trim(), @"\s{2,}");
            }

            if (parts.Length < 2) return null;

            // Try to find item number (first numeric field)
            int startIdx = 0;
            int itemNum = fallbackItemNum;
            if (int.TryParse(parts[0].Trim(), out int parsed))
            {
                itemNum = parsed;
                startIdx = 1;
            }

            if (startIdx >= parts.Length) return null;

            // Try to extract quantity (last numeric field)
            int qty = 1;
            int endIdx = parts.Length;
            if (endIdx > startIdx + 1 && int.TryParse(parts[endIdx - 1].Trim(), out int parsedQty) && parsedQty > 0 && parsedQty < 1000)
            {
                qty = parsedQty;
                endIdx--;
            }

            string partNumber = parts[startIdx].Trim();
            string description = endIdx > startIdx + 1
                ? string.Join(" ", parts.Skip(startIdx + 1).Take(endIdx - startIdx - 1)).Trim()
                : "";

            // Skip if part number looks like a header or noise
            if (partNumber.Length < 2) return null;
            string upper = partNumber.ToUpperInvariant();
            if (upper == "ITEM" || upper == "PART" || upper == "NO" || upper == "NUMBER") return null;

            return new BomRow
            {
                ItemNumber = itemNum,
                PartNumber = partNumber,
                Description = description,
                Quantity = qty
            };
        }

        private void HighlightComponent(IModelDoc2 model, IAssemblyDoc assembly, RenameEntry entry)
        {
            if (entry == null || string.IsNullOrEmpty(entry.CurrentFilePath)) return;

            try
            {
                // Find and select the component by path
                object compsObj = assembly.GetComponents(false);
                var comps = compsObj as object[];
                if (comps == null) return;

                foreach (var o in comps)
                {
                    var comp = o as IComponent2;
                    if (comp == null) continue;

                    string path = SafeGet(() => comp.GetPathName());
                    if (string.Equals(path, entry.CurrentFilePath, StringComparison.OrdinalIgnoreCase))
                    {
                        comp.Select4(false, null, false);
                        model.ViewZoomToSelection();
                        break;
                    }
                }
            }
            catch
            {
                // Non-critical — swallow selection errors
            }
        }

        private static string BrowseForPdf(string modelPath)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "Select assembly drawing PDF";
                dlg.Filter = "PDF Files|*.pdf|All Files|*.*";
                dlg.InitialDirectory = Path.GetDirectoryName(modelPath) ?? "";
                return dlg.ShowDialog() == DialogResult.OK ? dlg.FileName : null;
            }
        }

        private static T SafeGet<T>(Func<T> f)
        {
            try { return f(); }
            catch { return default(T); }
        }
    }
}
