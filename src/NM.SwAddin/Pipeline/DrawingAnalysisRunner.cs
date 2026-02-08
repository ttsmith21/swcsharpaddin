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
using NM.Core.Writeback;
using NM.Core.Writeback.Models;
using NM.SwAddin.Properties;
using NM.SwAddin.UI;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.Pipeline
{
    /// <summary>
    /// Orchestrates the full AI drawing analysis pipeline:
    ///   1. Find companion PDF
    ///   2. Analyze PDF (text + optional AI vision)
    ///   3. Read current custom properties from model
    ///   4. Reconcile 3D model data vs 2D drawing data
    ///   5. Generate property suggestions
    ///   6. Show PropertyReviewWizard for user approval
    ///   7. Apply approved suggestions to the model
    /// </summary>
    public sealed class DrawingAnalysisRunner
    {
        private readonly ISldWorks _swApp;
        private readonly IDrawingVisionService _visionService;

        public DrawingAnalysisRunner(ISldWorks swApp, IDrawingVisionService visionService = null)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
            _visionService = visionService ?? new OfflineVisionService();
        }

        /// <summary>
        /// Runs the full analysis pipeline on the active document.
        /// Shows the wizard and applies approved suggestions.
        /// </summary>
        /// <returns>Result summary, or null if cancelled.</returns>
        public DrawingAnalysisResult RunOnActiveDocument()
        {
            const string proc = nameof(RunOnActiveDocument);
            ErrorHandler.PushCallStack(proc);
            try
            {
                // 1. Get active model
                var model = _swApp.ActiveDoc as IModelDoc2;
                if (model == null)
                {
                    MessageBox.Show("No active document. Open a part or assembly first.",
                        "AI Drawing Analysis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return null;
                }

                string modelPath = model.GetPathName();
                if (string.IsNullOrEmpty(modelPath))
                {
                    MessageBox.Show("Please save the document first.",
                        "AI Drawing Analysis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return null;
                }

                // 2. Find companion PDF
                var analyzer = new PdfDrawingAnalyzer(_visionService);
                string pdfPath = analyzer.FindCompanionPdf(modelPath);

                if (string.IsNullOrEmpty(pdfPath))
                {
                    var result = MessageBox.Show(
                        $"No companion PDF found for:\n{Path.GetFileName(modelPath)}\n\n" +
                        "Searched: same folder, Drawings/, PDF/, parent folder.\n\n" +
                        "Would you like to browse for a PDF?",
                        "AI Drawing Analysis", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        pdfPath = BrowseForPdf(modelPath);
                    }

                    if (string.IsNullOrEmpty(pdfPath))
                        return null;
                }

                // 3. Analyze PDF
                DrawingData drawingData;
                if (_visionService.IsAvailable)
                {
                    drawingData = analyzer.AnalyzeWithVision(pdfPath);
                }
                else
                {
                    drawingData = analyzer.Analyze(pdfPath);
                }

                if (drawingData == null)
                {
                    MessageBox.Show("PDF analysis returned no data.",
                        "AI Drawing Analysis", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }

                // 4. Read current custom properties
                var modelInfo = new ModelInfo();
                var propsService = new CustomPropertiesService();
                propsService.ReadIntoCache(model, modelInfo);

                var currentProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var name in modelInfo.CustomProperties.GetAllPropertyNames())
                {
                    var val = modelInfo.CustomProperties.GetPropertyValue(name);
                    currentProps[name] = val?.ToString() ?? string.Empty;
                }

                // 5. Reconcile
                var partData = BuildPartDataFromProperties(currentProps, modelPath);
                var reconciliation = new ReconciliationEngine().Reconcile(partData, drawingData);

                // 6. Generate suggestions
                bool isAssembly = model.GetType() == (int)SolidWorks.Interop.swconst.swDocumentTypes_e.swDocASSEMBLY;
                var suggestionService = new PropertySuggestionService();

                List<PropertySuggestion> suggestions;
                if (isAssembly)
                {
                    suggestions = suggestionService.GenerateAssemblySuggestions(reconciliation, currentProps);
                }
                else
                {
                    suggestions = suggestionService.GeneratePartSuggestions(reconciliation, currentProps);
                }

                if (suggestions.Count == 0 && !reconciliation.HasConflicts && !reconciliation.HasRenameSuggestion)
                {
                    MessageBox.Show(
                        "Analysis complete. No property changes suggested.\n\n" +
                        $"{reconciliation.Confirmations.Count} fields confirmed matching.",
                        "AI Drawing Analysis", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return new DrawingAnalysisResult { Summary = "No changes needed", Applied = 0 };
                }

                // 7. Validate rename if suggested
                FileRenameValidation renameValidation = null;
                if (reconciliation.HasRenameSuggestion)
                {
                    renameValidation = new FileRenameValidator().Validate(reconciliation.Rename);
                }

                // 8. Show wizard
                string analysisDesc = drawingData.AnalysisMethod == AnalysisMethod.TextOnly
                    ? "Text extraction only"
                    : drawingData.AnalysisMethod == AnalysisMethod.VisionAI
                        ? "AI Vision analysis"
                        : "Hybrid (text + AI vision)";

                var wizard = new PropertyReviewWizard(
                    suggestions,
                    reconciliation.Conflicts,
                    reconciliation.Rename,
                    renameValidation,
                    Path.GetFileName(modelPath),
                    Path.GetFileName(pdfPath),
                    $"{reconciliation.Summary} | {analysisDesc}");

                if (wizard.ShowDialog() != DialogResult.OK)
                    return null;

                // 9. Apply approved suggestions
                var executor = new PropertyWritebackExecutor();
                var writeResult = executor.ApplyToCache(wizard.ApprovedSuggestions, modelInfo.CustomProperties);

                // Apply conflict resolutions
                foreach (var kv in wizard.ConflictResolutions)
                {
                    string propName = MapConflictFieldToProperty(kv.Key);
                    if (propName != null)
                    {
                        executor.ApplySingle(propName, kv.Value, modelInfo.CustomProperties);
                        writeResult.Applied.Add(new WritebackEntry
                        {
                            PropertyName = propName,
                            NewValue = kv.Value,
                            Status = WritebackStatus.Applied
                        });
                    }
                }

                // 10. Flush to SolidWorks
                if (writeResult.Applied.Count > 0)
                {
                    bool writeOk = propsService.WritePending(model, modelInfo);
                    if (!writeOk)
                    {
                        MessageBox.Show(
                            $"Applied {writeResult.Applied.Count} properties to cache, " +
                            "but WritePending failed. Properties may not be saved.",
                            "AI Drawing Analysis", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                // 11. Show summary
                string summaryMsg = $"Applied {writeResult.Applied.Count} properties, " +
                    $"skipped {writeResult.Skipped.Count}";
                if (writeResult.Failed.Count > 0)
                    summaryMsg += $", {writeResult.Failed.Count} failed";
                if (wizard.RenameApproved)
                    summaryMsg += "\n\nFile rename approved — rename must be done manually via File > Save As.";

                MessageBox.Show(summaryMsg, "AI Drawing Analysis - Complete",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                return new DrawingAnalysisResult
                {
                    Summary = writeResult.Summary,
                    Applied = writeResult.Applied.Count,
                    Skipped = writeResult.Skipped.Count,
                    Failed = writeResult.Failed.Count,
                    RenameApproved = wizard.RenameApproved
                };
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, ex.Message, ex);
                MessageBox.Show($"Analysis failed: {ex.Message}",
                    "AI Drawing Analysis", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                ErrorHandler.PopCallStack();
            }
        }

        private static string BrowseForPdf(string modelPath)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "Select companion PDF drawing";
                dlg.Filter = "PDF Files|*.pdf|All Files|*.*";
                dlg.InitialDirectory = Path.GetDirectoryName(modelPath) ?? "";
                return dlg.ShowDialog() == DialogResult.OK ? dlg.FileName : null;
            }
        }

        private static PartData BuildPartDataFromProperties(
            IDictionary<string, string> props, string filePath)
        {
            var pd = new PartData { FilePath = filePath };

            // ReconciliationEngine reads PartNumber/Description/Revision from Extra dict
            if (props.TryGetValue("Description", out var desc) && !string.IsNullOrWhiteSpace(desc))
                pd.Extra["Description"] = desc;
            if (props.TryGetValue("Print", out var pn) && !string.IsNullOrWhiteSpace(pn))
                pd.Extra["PartNumber"] = pn;
            if (props.TryGetValue("Revision", out var rev) && !string.IsNullOrWhiteSpace(rev))
                pd.Extra["Revision"] = rev;
            if (props.TryGetValue("OptiMaterial", out var mat) && !string.IsNullOrWhiteSpace(mat))
                pd.Material = mat;

            return pd;
        }

        private static string MapConflictFieldToProperty(string field)
        {
            switch (field)
            {
                case "PartNumber": return "Print";
                case "Description": return "Description";
                case "Revision": return "Revision";
                case "Material": return "OptiMaterial";
                default: return null;
            }
        }
    }

    /// <summary>
    /// Summary of a drawing analysis run.
    /// </summary>
    public sealed class DrawingAnalysisResult
    {
        public string Summary { get; set; }
        public int Applied { get; set; }
        public int Skipped { get; set; }
        public int Failed { get; set; }
        public bool RenameApproved { get; set; }
    }
}
