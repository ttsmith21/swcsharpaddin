using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NM.Core;
using NM.Core.DataModel;
using NM.Core.Export;
using NM.Core.Models;
using NM.Core.ProblemParts;
using NM.SwAddin.AssemblyProcessing;
using NM.SwAddin.Drawing;
using NM.SwAddin.UI;
using NM.SwAddin.Validation;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Pipeline
{
    /// <summary>
    /// Central orchestrator for the unified two-pass workflow.
    /// Detects context (part/assembly/drawing/folder) and runs:
    ///   Pass 1: Validate all models
    ///   Pass 2: Process good models (after user reviews problems)
    /// </summary>
    public sealed class WorkflowDispatcher
    {
        private readonly ISldWorks _swApp;
        private readonly BatchValidator _validator;

        public WorkflowDispatcher(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
            _validator = new BatchValidator(swApp);
        }

        /// <summary>
        /// Runs the unified workflow: detect context → validate all → show problems → process good.
        /// </summary>
        /// <param name="options">Processing options (optional).</param>
        /// <returns>WorkflowContext with results of both passes.</returns>
        public WorkflowContext Run(ProcessingOptions options = null)
        {
            const string proc = nameof(WorkflowDispatcher) + ".Run";
            ErrorHandler.PushCallStack(proc);

            options = options ?? new ProcessingOptions();
            var context = new WorkflowContext();
            context.StartTiming();

            try
            {
                // Step 1: Detect what's open and collect models
                DetectContext(context);
                if (context.UserCanceled)
                {
                    context.StopTiming();
                    return context;
                }

                CollectModels(context);
                if (context.AllModels.Count == 0)
                {
                    ShowMessage("No models found to process.", "Workflow");
                    context.StopTiming();
                    return context;
                }

                // Clear previous problem tracking
                ClearProblemManager();
                ValidationStats.Clear();

                // Step 2: PASS 1 - Validate all models
                using (var progressForm = new ProgressForm())
                {
                    progressForm.Text = "Validating Parts";
                    progressForm.SetMax(context.AllModels.Count);
                    progressForm.Show();
                    Application.DoEvents();

                    var validationResult = _validator.ValidateAll(
                        context.AllModels,
                        (current, total, fileName) =>
                        {
                            progressForm.SetStep(current, $"Validating: {fileName}");
                            Application.DoEvents();
                            if (progressForm.IsCanceled)
                                context.UserCanceled = true;
                        });

                    progressForm.Close();

                    if (context.UserCanceled)
                    {
                        context.StopTiming();
                        return context;
                    }

                    // Transfer results to context
                    context.GoodModels.AddRange(validationResult.GoodModels);
                    context.ProblemModels.AddRange(validationResult.ProblemModels);
                    context.ValidationElapsed = validationResult.Elapsed;
                    context.ValidationComplete = true;
                }

                // Step 3: Show problems if any
                if (context.ProblemModels.Count > 0)
                {
                    var action = ShowProblemPartsDialog(context);

                    switch (action)
                    {
                        case ProblemAction.Cancel:
                            context.UserCanceled = true;
                            context.StopTiming();
                            return context;

                        case ProblemAction.Retry:
                            // User may have fixed some - re-validate problem models
                            var nowGood = _validator.RevalidateModels(context.ProblemModels);
                            foreach (var model in nowGood)
                            {
                                context.ProblemModels.Remove(model);
                                context.GoodModels.Add(model);
                            }
                            break;

                        case ProblemAction.ContinueWithGood:
                            // Proceed with good models only
                            break;
                    }
                }

                // Step 4: PASS 2 - Process good models
                if (context.GoodModels.Count > 0)
                {
                    ProcessGoodModelsPass2(context, options);
                    AggregateCosts(context);  // Aggregate costs after processing
                    GenerateErpExport(context);  // Generate Import.prn if enabled
                }

                context.ProcessingComplete = true;

                // Step 5: Show final summary
                ShowFinalSummary(context);
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Workflow failed", ex, ErrorHandler.LogLevel.Error);
                ShowMessage($"Workflow error: {ex.Message}", "Error");
            }
            finally
            {
                context.StopTiming();
                ErrorHandler.PopCallStack();
            }

            return context;
        }

        private void ProcessGoodModelsPass2(WorkflowContext context, ProcessingOptions options)
        {
            using (var progressForm = new ProgressForm())
            {
                progressForm.Text = "Processing Parts";
                progressForm.SetMax(context.GoodModels.Count);
                progressForm.Show();
                Application.DoEvents();

                var sw = System.Diagnostics.Stopwatch.StartNew();
                int processed = 0;

                foreach (var modelInfo in context.GoodModels)
                {
                    processed++;
                    progressForm.SetStep(processed, $"Processing: {modelInfo.FileName}");
                    Application.DoEvents();

                    if (progressForm.IsCanceled)
                    {
                        context.UserCanceled = true;
                        break;
                    }

                    try
                    {
                        modelInfo.StartProcessing();
                        var partData = MainRunner.RunSinglePartData(_swApp, modelInfo.ModelDoc as IModelDoc2, options);
                        modelInfo.ProcessingResult = partData;  // Store for cost aggregation

                        var success = (partData?.Status == ProcessingStatus.Success);
                        if (success)
                        {
                            modelInfo.CompleteProcessing(true);
                            context.ProcessedModels.Add(modelInfo);
                        }
                        else
                        {
                            modelInfo.CompleteProcessing(false, partData?.FailureReason ?? "Processing failed");
                            context.FailedModels.Add(modelInfo);
                        }
                    }
                    catch (Exception ex)
                    {
                        modelInfo.CompleteProcessing(false, ex.Message);
                        context.FailedModels.Add(modelInfo);
                    }
                }

                sw.Stop();
                context.ProcessingElapsed = sw.Elapsed;
                progressForm.Close();
            }
        }

        private void AggregateCosts(WorkflowContext context)
        {
            double totalMaterial = 0;
            double totalProcessing = 0;
            double grandTotal = 0;

            foreach (var mi in context.ProcessedModels)
            {
                var cost = mi.ProcessingResult?.Cost;
                if (cost == null) continue;

                int qty = mi.Quantity;
                totalMaterial += cost.TotalMaterialCost * qty;
                totalProcessing += cost.TotalProcessingCost * qty;
                grandTotal += cost.TotalCost * qty;

                ErrorHandler.DebugLog($"[COST] {mi.FileName} x{qty}: ${cost.TotalCost:N2} ea = ${cost.TotalCost * qty:N2} ext");
            }

            context.TotalMaterialCost = totalMaterial;
            context.TotalProcessingCost = totalProcessing;
            context.GrandTotalCost = grandTotal;

            ErrorHandler.DebugLog($"[COST] Assembly Total: Material=${totalMaterial:N2}, Processing=${totalProcessing:N2}, TOTAL=${grandTotal:N2}");
        }

        private void GenerateErpExport(WorkflowContext context)
        {
            if (!context.GenerateErpExport || context.ProcessedModels.Count == 0)
                return;

            try
            {
                // Build PartData list from ProcessingResults
                var partDataList = context.ProcessedModels
                    .Where(m => m.ProcessingResult != null)
                    .Select(m => m.ProcessingResult)
                    .ToList();

                if (partDataList.Count == 0)
                {
                    ErrorHandler.DebugLog("[EXPORT] No parts with ProcessingResult - skipping export");
                    return;
                }

                // Determine parent part number from source
                string parentNumber = context.Source == WorkflowContext.SourceType.Assembly
                    ? Path.GetFileNameWithoutExtension(context.RootPath)
                    : "";

                var erpData = ErpExportDataBuilder.FromPartDataCollection(
                    partDataList,
                    parentNumber,
                    context.Customer);

                // Add BOM relationships from SwModelInfo quantities
                if (context.Source == WorkflowContext.SourceType.Assembly)
                {
                    erpData.BomRelationships.Clear();
                    int pieceNo = 1;
                    foreach (var mi in context.ProcessedModels)
                    {
                        if (mi.ProcessingResult == null) continue;
                        erpData.BomRelationships.Add(new BomRelationship
                        {
                            ParentPartNumber = parentNumber,
                            ChildPartNumber = Path.GetFileNameWithoutExtension(mi.FilePath),
                            PieceNumber = pieceNo.ToString(),
                            Quantity = mi.Quantity
                        });
                        pieceNo++;
                    }
                }

                // Export to Import.prn
                string exportPath = Path.Combine(
                    Path.GetDirectoryName(context.RootPath) ?? ".",
                    "Import.prn");

                var exporter = new ErpExportFormat { Customer = context.Customer };
                exporter.ExportToImportPrn(erpData, exportPath);
                context.ErpExportPath = exportPath;

                ErrorHandler.DebugLog($"[EXPORT] Generated: {exportPath}");
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError("GenerateErpExport", "Export failed", ex, ErrorHandler.LogLevel.Warning);
            }
        }

        private void ClearProblemManager()
        {
            // ProblemPartManager doesn't have Clear(), so remove items one by one
            var items = ProblemPartManager.Instance.GetProblemParts();
            foreach (var item in items.ToList())
            {
                ProblemPartManager.Instance.RemoveResolvedPart(item);
            }
        }

        private void DetectContext(WorkflowContext context)
        {
            var activeDoc = _swApp.ActiveDoc as IModelDoc2;

            if (activeDoc == null)
            {
                // No document open - prompt for folder
                context.Source = WorkflowContext.SourceType.Folder;
                context.RootPath = PromptForFolder();
                if (string.IsNullOrEmpty(context.RootPath))
                {
                    context.UserCanceled = true;
                }
            }
            else
            {
                var docType = (swDocumentTypes_e)activeDoc.GetType();
                context.SourceDocument = activeDoc;
                context.RootPath = activeDoc.GetPathName();

                switch (docType)
                {
                    case swDocumentTypes_e.swDocPART:
                        context.Source = WorkflowContext.SourceType.Part;
                        break;
                    case swDocumentTypes_e.swDocASSEMBLY:
                        context.Source = WorkflowContext.SourceType.Assembly;
                        break;
                    case swDocumentTypes_e.swDocDRAWING:
                        context.Source = WorkflowContext.SourceType.Drawing;
                        break;
                    default:
                        context.Source = WorkflowContext.SourceType.None;
                        break;
                }
            }

            ErrorHandler.DebugLog($"DetectContext: Source={context.Source}, Path={context.RootPath}");
        }

        private void CollectModels(WorkflowContext context)
        {
            switch (context.Source)
            {
                case WorkflowContext.SourceType.Part:
                    CollectSinglePart(context);
                    break;

                case WorkflowContext.SourceType.Assembly:
                    CollectAssemblyComponents(context);
                    break;

                case WorkflowContext.SourceType.Drawing:
                    CollectDrawingReferences(context);
                    break;

                case WorkflowContext.SourceType.Folder:
                    CollectFolderParts(context);
                    break;
            }

            ErrorHandler.DebugLog($"CollectModels: Found {context.AllModels.Count} models");
        }

        private void CollectSinglePart(WorkflowContext context)
        {
            var doc = context.SourceDocument;
            if (doc == null) return;

            var config = doc.ConfigurationManager?.ActiveConfiguration;
            string configName = config?.Name ?? string.Empty;

            var modelInfo = new SwModelInfo(doc.GetPathName(), configName)
            {
                ModelDoc = doc
            };
            context.AllModels.Add(modelInfo);
        }

        private void CollectAssemblyComponents(WorkflowContext context)
        {
            var doc = context.SourceDocument;
            if (doc == null) return;

            var assyDoc = doc as IAssemblyDoc;
            if (assyDoc == null) return;

            // Step 1: Get BOM quantities FIRST (the authoritative source)
            var quantifier = new AssemblyComponentQuantifier();
            var quantities = quantifier.CollectQuantitiesHybrid(assyDoc, "");
            ErrorHandler.DebugLog($"[ASMQTY] BOM returned {quantities.Count} unique entries");

            // Step 2: Get validated unique components
            var collector = new ComponentCollector();
            var result = collector.CollectUniqueComponents(assyDoc);

            // Step 3: Merge quantities into SwModelInfo
            foreach (var mi in result.ValidComponents)
            {
                string key = AssemblyComponentQuantifier.BuildKey(mi.FilePath, mi.Configuration);
                if (quantities.TryGetValue(key, out var q))
                {
                    mi.Quantity = q.Quantity;
                    ErrorHandler.DebugLog($"[ASMQTY] {mi.FileName} x{mi.Quantity}");
                }
                else
                {
                    mi.Quantity = 1; // Fallback if not in BOM
                    ErrorHandler.DebugLog($"[ASMQTY] {mi.FileName} not in BOM, defaulting to qty=1");
                }
            }

            // Add valid components to AllModels (now with quantities)
            context.AllModels.AddRange(result.ValidComponents);

            // Problem components go directly to ProblemModels
            foreach (var problem in result.ProblemComponents)
            {
                problem.MarkValidated(false, problem.ProblemDescription);
                context.ProblemModels.Add(problem);
            }

            // Store total BOM quantity for summary
            context.TotalBomQuantity = 0;
            foreach (var q in quantities.Values)
            {
                context.TotalBomQuantity += q.Quantity;
            }
            ErrorHandler.DebugLog($"[ASMQTY] Total BOM quantity: {context.TotalBomQuantity}");
        }

        private void CollectDrawingReferences(WorkflowContext context)
        {
            var doc = context.SourceDocument;
            if (doc == null) return;

            var extractor = new DrawingReferenceExtractor(_swApp);
            var models = extractor.ExtractReferences(doc);
            context.AllModels.AddRange(models);
        }

        private void CollectFolderParts(WorkflowContext context)
        {
            if (string.IsNullOrWhiteSpace(context.RootPath)) return;
            if (!Directory.Exists(context.RootPath)) return;

            var partFiles = Directory.GetFiles(context.RootPath, "*.sldprt", SearchOption.TopDirectoryOnly);

            foreach (var partFile in partFiles)
            {
                var modelInfo = new SwModelInfo(partFile);
                context.AllModels.Add(modelInfo);
            }
        }

        private string PromptForFolder()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select folder containing parts to process";
                dialog.ShowNewFolderButton = false;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.SelectedPath;
                }
            }
            return null;
        }

        private ProblemAction ShowProblemPartsDialog(WorkflowContext context)
        {
            using (var form = new ProblemPartsForm())
            {
                form.ShowGoodCount(context.GoodModels.Count);
                var result = form.ShowDialog();

                if (result == DialogResult.Cancel)
                    return ProblemAction.Cancel;

                return form.SelectedAction;
            }
        }

        private void ShowFinalSummary(WorkflowContext context)
        {
            string summary = context.GetFinalSummary();
            MessageBox.Show(summary, "Workflow Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowMessage(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>
    /// User action from ProblemPartsForm.
    /// </summary>
    public enum ProblemAction
    {
        Cancel = 0,
        Retry = 1,
        ContinueWithGood = 2
    }
}
