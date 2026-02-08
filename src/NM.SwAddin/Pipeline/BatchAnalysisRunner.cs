using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NM.Core;
using NM.Core.AI;
using NM.Core.DataModel;
using NM.Core.Pdf;
using NM.Core.Pdf.Models;
using NM.Core.Reconciliation;
using NM.Core.Reconciliation.Models;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Pipeline
{
    /// <summary>
    /// Orchestrates batch drawing analysis across an entire assembly and its drawing package.
    /// Flow:
    ///   1. User selects a folder of PDFs (or we auto-detect from assembly location)
    ///   2. DrawingPackageScanner builds an index of all pages by part number
    ///   3. Walk the assembly tree to collect unique components
    ///   4. ComponentDrawingMatcher matches each component to drawing pages
    ///   5. Run reconciliation + suggestion generation for each matched component
    ///   6. Present results for review
    /// </summary>
    public sealed class BatchAnalysisRunner
    {
        private readonly ISldWorks _swApp;
        private readonly IDrawingVisionService _visionService;

        public BatchAnalysisRunner(ISldWorks swApp, IDrawingVisionService visionService = null)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
            _visionService = visionService ?? new OfflineVisionService();
        }

        /// <summary>
        /// Runs batch analysis on the active assembly document.
        /// </summary>
        /// <returns>Batch result, or null if cancelled or no assembly open.</returns>
        public BatchAnalysisResult RunOnActiveAssembly()
        {
            const string proc = nameof(RunOnActiveAssembly);
            ErrorHandler.PushCallStack(proc);
            try
            {
                // 1. Validate active document is an assembly
                var model = _swApp.ActiveDoc as IModelDoc2;
                if (model == null)
                {
                    MessageBox.Show("No active document. Open an assembly first.",
                        "Batch Drawing Analysis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return null;
                }

                int docType = model.GetType();
                if (docType != (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    MessageBox.Show("Batch analysis requires an assembly.\nUse single-part analysis for individual parts.",
                        "Batch Drawing Analysis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return null;
                }

                string modelPath = model.GetPathName();
                if (string.IsNullOrEmpty(modelPath))
                {
                    MessageBox.Show("Please save the assembly first.",
                        "Batch Drawing Analysis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return null;
                }

                // 2. Locate drawing package folder
                string drawingFolder = FindDrawingPackageFolder(modelPath);
                if (string.IsNullOrEmpty(drawingFolder))
                {
                    var result = MessageBox.Show(
                        "No drawing package folder found.\n\n" +
                        "Searched: same folder, Drawings/, PDF/, parent folder.\n\n" +
                        "Would you like to browse for a folder?",
                        "Batch Drawing Analysis", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                        drawingFolder = BrowseForFolder(modelPath);

                    if (string.IsNullOrEmpty(drawingFolder))
                        return null;
                }

                // 3. Scan drawing package
                var scanner = new DrawingPackageScanner();
                var packageIndex = scanner.ScanFolder(drawingFolder);

                if (packageIndex.TotalPages == 0)
                {
                    MessageBox.Show($"No pages found in drawing package:\n{drawingFolder}",
                        "Batch Drawing Analysis", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }

                // Also scan single PDFs in the assembly's own directory if different
                string assemblyDir = Path.GetDirectoryName(modelPath);
                if (!string.Equals(assemblyDir, drawingFolder, StringComparison.OrdinalIgnoreCase))
                {
                    var localPdfs = Directory.GetFiles(assemblyDir, "*.pdf", SearchOption.TopDirectoryOnly);
                    foreach (var pdf in localPdfs)
                    {
                        if (!packageIndex.ScannedFiles.Contains(pdf))
                            scanner.ScanSinglePdf(pdf, packageIndex);
                    }
                }

                // 4. Collect assembly components
                var components = CollectComponents(model);

                // 5. Match components to drawing pages
                var matcher = new ComponentDrawingMatcher();
                var matchResults = matcher.MatchAll(components, packageIndex);

                // 6. Build batch result with per-component analysis
                var batchResult = new BatchAnalysisResult
                {
                    TopLevelName = Path.GetFileName(modelPath),
                    PackageIndex = packageIndex,
                    TotalComponents = components.Count
                };

                batchResult.UnmatchedDrawings.AddRange(matchResults.UnmatchedDrawings);

                foreach (string unmatched in matchResults.Unmatched)
                {
                    batchResult.UnmatchedComponents.Add(unmatched);
                }

                // For each matched component, run analysis
                foreach (var kv in matchResults.Matched)
                {
                    string compPath = kv.Key;
                    var match = kv.Value;
                    var comp = components.FirstOrDefault(c =>
                        string.Equals(c.FilePath, compPath, StringComparison.OrdinalIgnoreCase));

                    if (comp == null) continue;

                    var drawingData = packageIndex.BuildDrawingData(
                        match.Pages[0].PartNumber ?? Path.GetFileNameWithoutExtension(compPath));

                    var compResult = new ComponentAnalysisResult
                    {
                        FilePath = compPath,
                        FileName = Path.GetFileNameWithoutExtension(compPath),
                        PartNumber = comp.PartNumber,
                        IsAssembly = comp.IsAssembly,
                        Quantity = comp.Quantity,
                        DrawingData = drawingData,
                        MatchConfidence = match.Confidence,
                        MatchMethod = match.Method
                    };

                    // Run reconciliation if we have drawing data
                    if (drawingData != null)
                    {
                        var partData = new PartData { FilePath = compPath };
                        if (!string.IsNullOrEmpty(comp.PartNumber))
                            partData.Extra["PartNumber"] = comp.PartNumber;

                        var reconciliation = new ReconciliationEngine().Reconcile(partData, drawingData);
                        var suggestionService = new PropertySuggestionService();

                        List<PropertySuggestion> suggestions;
                        if (comp.IsAssembly)
                        {
                            suggestions = suggestionService.GenerateAssemblySuggestions(
                                reconciliation, new Dictionary<string, string>());
                        }
                        else
                        {
                            suggestions = suggestionService.GeneratePartSuggestions(
                                reconciliation, new Dictionary<string, string>());
                        }

                        compResult.SuggestionCount = suggestions.Count;
                        compResult.ConflictCount = reconciliation.Conflicts.Count;
                        compResult.HasRenameSuggestion = reconciliation.HasRenameSuggestion;
                    }

                    batchResult.ComponentResults[compPath] = compResult;
                }

                // 7. Show summary
                string summaryMsg =
                    $"Drawing Package: {packageIndex.Summary}\n\n" +
                    $"Assembly: {batchResult.TopLevelName}\n" +
                    $"Components: {batchResult.TotalComponents}\n" +
                    $"Matched: {batchResult.MatchedComponents}\n" +
                    $"Unmatched parts: {batchResult.UnmatchedComponents.Count}\n" +
                    $"Unmatched drawings: {batchResult.UnmatchedDrawings.Count}\n" +
                    $"Total suggestions: {batchResult.TotalSuggestions}";

                MessageBox.Show(summaryMsg, "Batch Drawing Analysis - Complete",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                return batchResult;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                MessageBox.Show($"Batch analysis failed: {ex.Message}",
                    "Batch Drawing Analysis", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        /// <summary>
        /// Collects unique components from the assembly, extracting part numbers
        /// from custom properties where available.
        /// </summary>
        private List<ComponentInfo> CollectComponents(IModelDoc2 model)
        {
            var result = new List<ComponentInfo>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var asm = model as IAssemblyDoc;
            if (asm == null) return result;

            var config = model.ConfigurationManager?.ActiveConfiguration;
            var root = config?.GetRootComponent3(true) as IComponent2;
            if (root == null) return result;

            CollectComponentsRecursive(root, result, seen);
            return result;
        }

        private void CollectComponentsRecursive(
            IComponent2 comp, List<ComponentInfo> result, HashSet<string> seen)
        {
            var kids = comp.GetChildren();
            if (!(kids is Array arr)) return;

            foreach (var childObj in arr)
            {
                var child = childObj as IComponent2;
                if (child == null) continue;

                // Skip suppressed
                try
                {
                    if (child.GetSuppression2() == (int)swComponentSuppressionState_e.swComponentSuppressed)
                        continue;
                }
                catch { continue; }

                string childPath;
                try { childPath = child.GetPathName(); }
                catch { continue; }

                if (string.IsNullOrEmpty(childPath))
                    continue;

                bool isAssy = childPath.EndsWith(".SLDASM", StringComparison.OrdinalIgnoreCase);

                if (seen.Add(childPath))
                {
                    string partNumber = null;
                    try
                    {
                        var childModel = child.GetModelDoc2() as IModelDoc2;
                        if (childModel != null)
                        {
                            var propMgr = childModel.Extension?.CustomPropertyManager[""];
                            if (propMgr != null)
                            {
                                string valOut, resolvedOut;
                                bool wasResolved;
                                propMgr.Get5("Print", true, out valOut, out resolvedOut, out wasResolved);
                                partNumber = resolvedOut;
                            }
                        }
                    }
                    catch { /* Lightweight or unresolved component */ }

                    result.Add(new ComponentInfo
                    {
                        FilePath = childPath,
                        PartNumber = partNumber,
                        IsAssembly = isAssy,
                        Quantity = 1
                    });
                }
                else
                {
                    // Already seen — increment quantity
                    var existing = result.FirstOrDefault(c =>
                        string.Equals(c.FilePath, childPath, StringComparison.OrdinalIgnoreCase));
                    if (existing != null)
                        existing.Quantity++;
                }

                // Recurse into sub-assemblies
                if (isAssy)
                {
                    CollectComponentsRecursive(child, result, seen);
                }
            }
        }

        /// <summary>
        /// Finds the drawing package folder near the assembly.
        /// </summary>
        private static string FindDrawingPackageFolder(string assemblyPath)
        {
            string dir = Path.GetDirectoryName(assemblyPath);
            if (dir == null) return null;

            // Check if there are PDFs in the current directory
            if (Directory.GetFiles(dir, "*.pdf", SearchOption.TopDirectoryOnly).Length > 0)
                return dir;

            // Check common subfolder names
            string[] subfolderNames = { "Drawings", "PDF", "PDFs", "Drawing Package", "Dwgs" };
            foreach (var name in subfolderNames)
            {
                string sub = Path.Combine(dir, name);
                if (Directory.Exists(sub) &&
                    Directory.GetFiles(sub, "*.pdf", SearchOption.TopDirectoryOnly).Length > 0)
                    return sub;
            }

            // Check parent directory
            string parent = Directory.GetParent(dir)?.FullName;
            if (parent != null)
            {
                if (Directory.GetFiles(parent, "*.pdf", SearchOption.TopDirectoryOnly).Length > 0)
                    return parent;

                foreach (var name in subfolderNames)
                {
                    string sub = Path.Combine(parent, name);
                    if (Directory.Exists(sub) &&
                        Directory.GetFiles(sub, "*.pdf", SearchOption.TopDirectoryOnly).Length > 0)
                        return sub;
                }
            }

            return null;
        }

        private static string BrowseForFolder(string assemblyPath)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select folder containing PDF drawing package";
                dlg.SelectedPath = Path.GetDirectoryName(assemblyPath) ?? "";
                return dlg.ShowDialog() == DialogResult.OK ? dlg.SelectedPath : null;
            }
        }
    }
}
