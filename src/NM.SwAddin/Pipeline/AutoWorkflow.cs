using System;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NM.Core;
using NM.Core.DataModel;
using NM.Core.ProblemParts;
using NM.SwAddin.AssemblyProcessing;
using NM.SwAddin.Processing;
using NM.SwAddin.UI;
using NM.SwAddin.Utils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Pipeline
{
    /// <summary>
    /// Auto-dispatcher: assembly ? assembly workflow; part ? single-part; no doc ? folder selection + batch.
    /// </summary>
    public static class AutoWorkflow
    {
        public static void Run(ISldWorks swApp, bool? forceVisible = null)
        {
            if (swApp == null)
            {
                MessageBox.Show("No SolidWorks application handle.");
                return;
            }

            // Default: hide SW for speed unless explicitly overridden
            bool targetVisible = forceVisible ?? false;
            using (new VisibilityScope(swApp, targetVisible))
            {
                var doc = swApp.ActiveDoc as IModelDoc2;
                if (doc == null)
                {
                    RunFolderSelection(swApp);
                    return;
                }

                var type = (swDocumentTypes_e)doc.GetType();
                switch (type)
                {
                    case swDocumentTypes_e.swDocASSEMBLY:
                        RunAssembly(swApp, (IAssemblyDoc)doc);
                        break;

                    case swDocumentTypes_e.swDocPART:
                        RunSinglePart(swApp, doc);
                        break;

                    default:
                        swApp.SendMsgToUser2("Unsupported active document type. Choose a folder to batch process instead.",
                            (int)swMessageBoxIcon_e.swMbInformation, (int)swMessageBoxBtn_e.swMbOk);
                        RunFolderSelection(swApp);
                        break;
                }
            }
        }

        private static void RunSinglePart(ISldWorks swApp, IModelDoc2 partDoc)
        {
            var result = MainRunner.RunSinglePart(swApp, partDoc, options: null);
            swApp.SendMsgToUser2(
                result.Success ? $"Part OK: {result.Message}" : $"Part failed: {result.Message}",
                result.Success ? (int)swMessageBoxIcon_e.swMbInformation : (int)swMessageBoxIcon_e.swMbStop,
                (int)swMessageBoxBtn_e.swMbOk);
        }

        private static void RunAssembly(ISldWorks swApp, IAssemblyDoc asm)
        {
            try
            {
                var asmModel = (IModelDoc2)asm;
                var pre = new AssemblyPreprocessor();
                pre.EnsureComponentsResolved(asm);

                // Get BOM quantities
                var quantifier = new AssemblyComponentQuantifier();
                var quantities = quantifier.CollectQuantitiesHybrid(asm, "");

                var prep = pre.PreprocessAssembly(asmModel);
                var sb = new StringBuilder();
                sb.AppendLine(prep.Summary);

                if (prep.ComponentsToProcess == null || prep.ComponentsToProcess.Count == 0)
                {
                    swApp.SendMsgToUser2(
                        "Assembly detected but no valid unique components to process.\nReview Problem Parts for details.",
                        (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
                    TryShowProblemPartsUI();
                    return;
                }

                int ok = 0, fail = 0;
                double totalMaterial = 0, totalProcessing = 0, grandTotal = 0;
                int totalBomQty = 0;
                var fileOps = new SolidWorksFileOperations(swApp);

                // Progress dialog so the user knows SolidWorks isn't frozen
                // (graphics updates are suppressed inside BatchPerformanceScope)
                using (var progressForm = new ProgressForm())
                {
                    progressForm.Text = "Processing Assembly Components";
                    progressForm.SetMax(prep.ComponentsToProcess.Count);
                    progressForm.Show();
                    System.Windows.Forms.Application.DoEvents();

                    int step = 0;

                // Batch performance optimization: disable graphics updates during component loop
                using (new BatchPerformanceScope(swApp, asmModel))
                foreach (var mi in prep.ComponentsToProcess)
                {
                    step++;
                    progressForm.SetStep(step, $"Processing: {mi.FileName}");
                    System.Windows.Forms.Application.DoEvents();

                    if (progressForm.IsCanceled) break;

                    // Get quantity from BOM
                    string key = AssemblyComponentQuantifier.BuildKey(mi.FilePath, mi.Configuration);
                    int qty = quantities.TryGetValue(key, out var q) ? q.Quantity : 1;
                    totalBomQty += qty;

                    IModelDoc2 compDoc = null;
                    try
                    {
                        compDoc = fileOps.OpenSWDocument(mi.FilePath, silent: true, readOnly: false, configurationName: mi.Configuration);
                        if (compDoc == null)
                        {
                            fail++;
                            ProblemPartManager.Instance.AddProblemPart(
                                new NM.Core.ModelInfo(), "Failed to open component", ProblemPartManager.ProblemCategory.FileAccess);
                            continue;
                        }

                        var partData = MainRunner.RunSinglePartData(swApp, compDoc, options: null);
                        if (partData?.Status == ProcessingStatus.Success)
                        {
                            ok++;
                            var cost = partData.Cost;
                            if (cost != null)
                            {
                                totalMaterial += cost.TotalMaterialCost * qty;
                                totalProcessing += cost.TotalProcessingCost * qty;
                                grandTotal += cost.TotalCost * qty;
                            }
                        }
                        else
                        {
                            fail++;
                            ProblemPartManager.Instance.AddProblemPart(
                                new NM.Core.ModelInfo(), partData?.FailureReason ?? "Processing failed", ProblemPartManager.ProblemCategory.Fatal);
                        }
                    }
                    catch (Exception ex)
                    {
                        fail++;
                        ProblemPartManager.Instance.AddProblemPart(
                            new NM.Core.ModelInfo(), ex.Message, ProblemPartManager.ProblemCategory.Fatal);
                    }
                    finally
                    {
                        try { if (compDoc != null) fileOps.CloseSWDocument(compDoc); } catch { }
                    }
                }

                } // close progressForm using block

                sb.AppendLine();
                sb.AppendLine($"Processed OK: {ok}");
                sb.AppendLine($"Failed: {fail}");
                sb.AppendLine($"Total BOM Quantity: {totalBomQty}");
                if (grandTotal > 0)
                {
                    sb.AppendLine();
                    sb.AppendLine("Cost Summary:");
                    sb.AppendLine($"  Material: ${totalMaterial:N2}");
                    sb.AppendLine($"  Processing: ${totalProcessing:N2}");
                    sb.AppendLine($"  TOTAL: ${grandTotal:N2}");
                }

                swApp.SendMsgToUser2(sb.ToString(),
                    (fail == 0) ? (int)swMessageBoxIcon_e.swMbInformation : (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk);

                if (prep.ProblemParts != null && prep.ProblemParts.Count > 0)
                {
                    TryShowProblemPartsUI();
                }
            }
            catch (Exception ex)
            {
                swApp.SendMsgToUser2("Assembly workflow error: " + ex.Message,
                    (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
            }
        }

        private static void RunFolderSelection(ISldWorks swApp)
        {
            try
            {
                using (var dlg = new FolderBrowserDialog())
                {
                    dlg.Description = "Select a root folder to process";
                    if (dlg.ShowDialog() != DialogResult.OK) return;

                    var processor = new FolderProcessor(swApp);
                    var res = processor.ProcessFolder(dlg.SelectedPath, recursive: true);

                    var sb = new StringBuilder();
                    sb.AppendLine("Folder Processing Summary:");
                    sb.AppendLine($"Root: {dlg.SelectedPath}");
                    sb.AppendLine($"Discovered: {res.TotalDiscovered}");
                    sb.AppendLine($"Imported: {res.ImportedCount}");
                    sb.AppendLine($"Opened: {res.OpenedOk}");
                    sb.AppendLine($"Processed: {res.Processed}");
                    sb.AppendLine($"Failed: {res.FailedOpen}");
                    sb.AppendLine($"Skipped: {res.Skipped}");
                    if (res.Errors.Count > 0)
                    {
                        sb.AppendLine();
                        sb.AppendLine("Errors:");
                        foreach (var e in res.Errors.Take(25)) sb.AppendLine(" - " + e);
                        if (res.Errors.Count > 25) sb.AppendLine($" - ... and {res.Errors.Count - 25} more");
                    }
                    MessageBox.Show(sb.ToString(), "Process Folder");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Folder processing error: " + ex.Message);
            }
        }

        private static void TryShowProblemPartsUI()
        {
            try
            {
                var form = new ProblemPartsForm(ProblemPartManager.Instance);
                form.StartPosition = FormStartPosition.CenterScreen;
                form.Show();
            }
            catch { }
        }
    }
}
