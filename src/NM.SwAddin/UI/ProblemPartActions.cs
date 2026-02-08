using System;
using System.Collections.Generic;
using System.Linq;
using NM.Core;
using NM.Core.DataModel;
using NM.Core.Models;
using NM.Core.ProblemParts;
using NM.SwAddin.Geometry;
using NM.SwAddin.Pipeline;
using NM.SwAddin.Processing;
using NM.SwAddin.Validation;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using static NM.Core.Constants.UnitConversions;

namespace NM.SwAddin.UI
{
    /// <summary>
    /// Result of a problem part action (classify, process, retry, split).
    /// </summary>
    public sealed class ActionResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public PartData ProcessingResult { get; set; }
        public List<PartData> SubPartResults { get; set; }
    }

    /// <summary>
    /// Shared business logic for problem part resolution.
    /// Used by both ProblemWizardForm (floating dialog) and ProblemPartsTaskPaneControl (docked panel).
    /// Pure operations — no UI code. Callers update their own controls based on ActionResult.
    /// </summary>
    public static class ProblemPartActions
    {
        /// <summary>
        /// Classifies a problem part as PUR/MACH/CUST by writing custom properties
        /// and marking it resolved in ProblemPartManager.
        /// </summary>
        public static ActionResult ClassifyPart(
            ProblemPartManager.ProblemItem item,
            ProblemPartManager.PartTypeOverride typeOverride,
            IModelDoc2 currentDoc,
            List<ProblemPartManager.ProblemItem> fixedProblems)
        {
            if (item == null)
                return new ActionResult { Success = false, Message = "No problem item specified." };

            ProblemPartManager.Instance.SetTypeOverride(item, typeOverride);

            if (currentDoc != null)
            {
                SwPropertyHelper.AddCustomProperty(currentDoc, "rbPartType",
                    swCustomInfoType_e.swCustomInfoNumber, "1", "");
                SwPropertyHelper.AddCustomProperty(currentDoc, "rbPartTypeSub",
                    swCustomInfoType_e.swCustomInfoNumber, ((int)typeOverride).ToString(), "");
                SwDocumentHelper.SaveDocument(currentDoc);
            }

            fixedProblems.Add(item);
            ProblemPartManager.Instance.RemoveResolvedPart(item);

            return new ActionResult
            {
                Success = true,
                Message = $"Classified as {typeOverride}: {item.DisplayName}"
            };
        }

        /// <summary>
        /// Runs the part through MainRunner with a forced classification (SheetMetal or Tube).
        /// </summary>
        public static ActionResult RunAsClassification(
            ProblemPartManager.ProblemItem item,
            string classification,
            ISldWorks swApp,
            IModelDoc2 currentDoc,
            List<ProblemPartManager.ProblemItem> fixedProblems)
        {
            if (currentDoc == null || item == null)
                return new ActionResult { Success = false, Message = "No document or problem item." };

            try
            {
                var options = new ProcessingOptions { ForceClassification = classification };
                var result = MainRunner.RunSinglePartData(swApp, currentDoc, options);

                if (result?.Status == ProcessingStatus.Success)
                {
                    SwDocumentHelper.SaveDocument(currentDoc);
                    item.Metadata["AlreadyProcessed"] = "true";
                    item.Metadata["ProcessingResult"] = result;
                    fixedProblems.Add(item);
                    ProblemPartManager.Instance.RemoveResolvedPart(item);

                    return new ActionResult
                    {
                        Success = true,
                        Message = $"Processed as {classification}: {item.DisplayName}",
                        ProcessingResult = result
                    };
                }

                return new ActionResult
                {
                    Success = false,
                    Message = $"Failed as {classification}: {result?.FailureReason}"
                };
            }
            catch (Exception ex)
            {
                return new ActionResult { Success = false, Message = $"Error: {ex.Message}" };
            }
        }

        /// <summary>
        /// Splits a multi-body part into individual parts + assembly, then processes each sub-part.
        /// </summary>
        public static ActionResult RunSplitToAssembly(
            ProblemPartManager.ProblemItem item,
            ISldWorks swApp,
            IModelDoc2 currentDoc,
            List<ProblemPartManager.ProblemItem> fixedProblems,
            Action<string> statusCallback = null)
        {
            if (currentDoc == null || item == null)
                return new ActionResult { Success = false, Message = "No document or problem item." };

            var partDoc = currentDoc as IPartDoc;
            if (partDoc == null)
                return new ActionResult { Success = false, Message = "Split requires a part document (not assembly/drawing)." };

            var bodiesObj = partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
            if (bodiesObj == null || ((object[])bodiesObj).Length < 2)
                return new ActionResult { Success = false, Message = "Part must have 2 or more solid bodies to split." };

            try
            {
                SwDocumentHelper.SaveDocument(currentDoc);

                var splitter = new MultiBodySplitter(swApp);
                var splitResult = splitter.SplitToAssembly(currentDoc);

                if (!splitResult.Success)
                    return new ActionResult { Success = false, Message = $"Split failed: {splitResult.ErrorMessage}" };

                statusCallback?.Invoke($"Split into {splitResult.BodyCount} parts. Processing sub-parts...");

                var subResults = new List<PartData>();
                int processed = 0;

                foreach (var subPartPath in splitResult.PartPaths)
                {
                    processed++;
                    statusCallback?.Invoke($"Processing sub-part {processed}/{splitResult.PartPaths.Length}: {System.IO.Path.GetFileName(subPartPath)}");

                    IModelDoc2 subDoc = null;
                    try
                    {
                        int errs = 0, warns = 0;
                        subDoc = swApp.OpenDoc6(subPartPath, (int)swDocumentTypes_e.swDocPART,
                            (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errs, ref warns) as IModelDoc2;

                        if (subDoc == null)
                        {
                            ErrorHandler.DebugLog($"[SPLIT] Failed to open sub-part: {subPartPath}");
                            continue;
                        }

                        PreparePartView(subDoc);
                        var partData = MainRunner.RunSinglePartData(swApp, subDoc, null);

                        if (partData?.Status == ProcessingStatus.Success)
                        {
                            SwDocumentHelper.SaveDocument(subDoc);
                            subResults.Add(partData);
                        }
                        else
                        {
                            ErrorHandler.DebugLog($"[SPLIT] Sub-part processing failed: {subPartPath} - {partData?.FailureReason}");
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorHandler.DebugLog($"[SPLIT] Sub-part error: {subPartPath} - {ex.Message}");
                    }
                    finally
                    {
                        if (subDoc != null)
                        {
                            try { swApp.CloseDoc(subDoc.GetTitle()); } catch { }
                        }
                    }
                }

                item.Metadata["AlreadyProcessed"] = "true";
                item.Metadata["SubPartResults"] = subResults;
                fixedProblems.Add(item);
                ProblemPartManager.Instance.RemoveResolvedPart(item);

                return new ActionResult
                {
                    Success = true,
                    Message = $"Split complete: {splitResult.BodyCount} bodies, {subResults.Count} processed. Assembly: {System.IO.Path.GetFileName(splitResult.AssemblyPath)}",
                    SubPartResults = subResults
                };
            }
            catch (Exception ex)
            {
                return new ActionResult { Success = false, Message = $"Split error: {ex.Message}" };
            }
        }

        /// <summary>
        /// Revalidates a problem part. If validation passes, marks it fixed.
        /// </summary>
        public static ActionResult RetryAndValidate(
            ProblemPartManager.ProblemItem item,
            ISldWorks swApp,
            IModelDoc2 currentDoc,
            List<ProblemPartManager.ProblemItem> fixedProblems)
        {
            if (swApp == null || item == null)
                return new ActionResult { Success = false, Message = "Missing SolidWorks instance or problem item." };

            try
            {
                var model = currentDoc;
                if (model == null || !string.Equals(model.GetPathName(), item.FilePath, StringComparison.OrdinalIgnoreCase))
                {
                    int errs = 0, warns = 0;
                    model = swApp.OpenDoc6(item.FilePath, SwDocumentHelper.GuessDocType(item.FilePath), 0, item.Configuration ?? "", ref errs, ref warns) as IModelDoc2;
                }

                if (model == null)
                    return new ActionResult { Success = false, Message = "Could not open part for validation." };

                var swInfo = new SwModelInfo(item.FilePath) { Configuration = item.Configuration ?? "" };
                var validator = new PartValidationAdapter();
                var vr = validator.Validate(swInfo, model);

                if (vr.Success)
                {
                    fixedProblems.Add(item);
                    ProblemPartManager.Instance.RemoveResolvedPart(item);
                    return new ActionResult { Success = true, Message = $"FIXED: {item.DisplayName}" };
                }

                item.ProblemDescription = vr.Summary;
                item.RetryCount++;
                return new ActionResult { Success = false, Message = $"Still failing: {vr.Summary}" };
            }
            catch (Exception ex)
            {
                return new ActionResult { Success = false, Message = $"Validation error: {ex.Message}" };
            }
        }

        /// <summary>
        /// Writes NM_SkippedReason custom property to the document.
        /// </summary>
        public static void MarkAsSkipped(ProblemPartManager.ProblemItem item, IModelDoc2 currentDoc)
        {
            if (currentDoc == null || item == null) return;

            try
            {
                string reason = item.ProblemDescription ?? "Unknown";
                string value = $"{reason} [Skipped {DateTime.Now:yyyy-MM-dd HH:mm}]";

                SwPropertyHelper.AddCustomProperty(
                    currentDoc,
                    "NM_SkippedReason",
                    swCustomInfoType_e.swCustomInfoText,
                    value,
                    "");

                ErrorHandler.DebugLog($"[Actions] Marked skipped: {item.DisplayName} - {reason}");
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[Actions] Failed to mark skipped: {ex.Message}");
            }
        }

        /// <summary>
        /// Opens a problem part in SolidWorks and returns the document.
        /// </summary>
        public static IModelDoc2 OpenPart(ProblemPartManager.ProblemItem item, ISldWorks swApp)
        {
            if (swApp == null || item == null || string.IsNullOrEmpty(item.FilePath))
                return null;

            try
            {
                int errs = 0, warns = 0;
                int docType = SwDocumentHelper.GuessDocType(item.FilePath);

                var doc = swApp.OpenDoc6(item.FilePath, docType, 0, item.Configuration ?? "", ref errs, ref warns) as IModelDoc2;
                if (doc == null || errs != 0)
                    return null;

                int activateErr = 0;
                swApp.ActivateDoc3(doc.GetTitle(), true, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref activateErr);

                PreparePartView(doc);
                return doc;
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[Actions] Open failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Saves and closes a document.
        /// </summary>
        public static void SaveAndClosePart(IModelDoc2 doc, ISldWorks swApp)
        {
            if (doc == null) return;

            try { SwDocumentHelper.SaveDocument(doc); }
            catch (Exception ex) { ErrorHandler.DebugLog($"[Actions] Save failed: {ex.Message}"); }

            try { SwDocumentHelper.CloseDocument(swApp, doc); }
            catch (Exception ex) { ErrorHandler.DebugLog($"[Actions] Close failed: {ex.Message}"); }
        }

        /// <summary>
        /// Extracts tube diagnostic info from the current document if not already cached.
        /// </summary>
        public static bool EnsureTubeDiagnostics(
            ISldWorks swApp, IModelDoc2 currentDoc,
            ref TubeDiagnosticInfo diagnostics, out string statusMessage)
        {
            statusMessage = "";

            if (swApp == null) { statusMessage = "SolidWorks instance not available."; return false; }
            if (currentDoc == null) { statusMessage = "Part is not open."; return false; }

            if (diagnostics == null)
            {
                try
                {
                    var extractor = new TubeGeometryExtractor(swApp);
                    var (profile, diag) = extractor.ExtractWithDiagnostics(currentDoc);
                    diagnostics = diag;

                    statusMessage = profile != null
                        ? $"Profile: {profile.Shape}, OD={profile.OuterDiameterMeters * MetersToInches:F3}in | {diag.GetSummary()}"
                        : "No tube profile detected. " + diag.GetSummary();
                }
                catch (Exception ex)
                {
                    statusMessage = "Extraction failed: " + ex.Message;
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Selects tube diagnostic geometry in SolidWorks for visual debugging.
        /// </summary>
        public static string SelectTubeDiagnostic(
            ISldWorks swApp, IModelDoc2 currentDoc,
            TubeDiagnosticInfo diagnostics, TubeDiagnosticKind kind)
        {
            if (currentDoc == null) return "No model open.";
            if (diagnostics == null) return "No diagnostics available.";

            try
            {
                var extractor = new TubeGeometryExtractor(swApp);
                switch (kind)
                {
                    case TubeDiagnosticKind.CutLength:
                        extractor.SelectCutLengthEdges(currentDoc, diagnostics);
                        return $"Selected {diagnostics.CutLengthEdges.Count} cut length edges (green).";
                    case TubeDiagnosticKind.Holes:
                        extractor.SelectHoleEdges(currentDoc, diagnostics);
                        return $"Selected {diagnostics.HoleEdges.Count} hole edges (red).";
                    case TubeDiagnosticKind.Boundary:
                        extractor.SelectBoundaryEdges(currentDoc, diagnostics);
                        return $"Selected {diagnostics.BoundaryEdges.Count} boundary edges (blue).";
                    case TubeDiagnosticKind.Profile:
                        extractor.SelectProfileFaces(currentDoc, diagnostics);
                        return $"Selected {diagnostics.ProfileFaces.Count} profile faces (cyan).";
                    case TubeDiagnosticKind.All:
                        extractor.SelectAllDiagnostics(currentDoc, diagnostics);
                        return diagnostics.GetSummary();
                    case TubeDiagnosticKind.Clear:
                        currentDoc.ClearSelection2(true);
                        return "Selection cleared.";
                    default:
                        return "Unknown diagnostic kind.";
                }
            }
            catch (Exception ex)
            {
                return "Select failed: " + ex.Message;
            }
        }

        /// <summary>
        /// Sets display mode to Shaded with Edges and zooms to fit.
        /// </summary>
        public static void PreparePartView(IModelDoc2 doc)
        {
            if (doc == null) return;
            try
            {
                var view = doc.ActiveView as IModelView;
                if (view != null)
                    view.DisplayMode = (int)swViewDisplayMode_e.swViewDisplayMode_ShadedWithEdges;
                doc.ViewZoomtofit2();
            }
            catch { }
        }
    }

    /// <summary>
    /// Kinds of tube diagnostic visualizations.
    /// </summary>
    public enum TubeDiagnosticKind
    {
        CutLength,
        Holes,
        Boundary,
        Profile,
        All,
        Clear
    }
}
