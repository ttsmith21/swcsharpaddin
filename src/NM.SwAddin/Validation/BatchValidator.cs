using System;
using System.Collections.Generic;
using System.Diagnostics;
using NM.Core;
using NM.Core.Models;
using NM.Core.ProblemParts;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace NM.SwAddin.Validation
{
    /// <summary>
    /// Pass 1: Validates all collected models before processing.
    /// Opens each file, runs preflight, categorizes as Good or Problem.
    /// </summary>
    public sealed class BatchValidator
    {
        private readonly ISldWorks _swApp;
        private readonly SolidWorksFileOperations _fileOps;
        private readonly PartValidationAdapter _validator;

        public BatchValidator(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
            _fileOps = new SolidWorksFileOperations(swApp);
            _validator = new PartValidationAdapter();
        }

        /// <summary>
        /// Result of batch validation.
        /// </summary>
        public sealed class BatchValidationResult
        {
            public List<SwModelInfo> GoodModels { get; } = new List<SwModelInfo>();
            public List<SwModelInfo> ProblemModels { get; } = new List<SwModelInfo>();
            public int TotalValidated { get; set; }
            public TimeSpan Elapsed { get; set; }

            public string GetSummary()
            {
                return $"Validated {TotalValidated}: {GoodModels.Count} OK, {ProblemModels.Count} problems ({Elapsed.TotalSeconds:F1}s)";
            }
        }

        /// <summary>
        /// Validates all models in the list, updating their ProcessingState.
        /// </summary>
        /// <param name="models">Models to validate.</param>
        /// <param name="progressCallback">Optional callback (currentIndex, totalCount, currentFileName).</param>
        /// <returns>Validation result with good and problem model lists.</returns>
        public BatchValidationResult ValidateAll(
            IList<SwModelInfo> models,
            Action<int, int, string> progressCallback = null)
        {
            const string proc = nameof(BatchValidator) + ".ValidateAll";
            ErrorHandler.PushCallStack(proc);

            var result = new BatchValidationResult();
            var sw = Stopwatch.StartNew();

            ValidationStats.Clear();

            try
            {
                if (models == null || models.Count == 0)
                {
                    result.Elapsed = sw.Elapsed;
                    return result;
                }

                for (int i = 0; i < models.Count; i++)
                {
                    var model = models[i];
                    if (model == null) continue;

                    ErrorHandler.DebugLog($"[BATCHVAL] Loop iteration {i + 1}/{models.Count}: {model.FileName}");

                    try
                    {
                        progressCallback?.Invoke(i + 1, models.Count, model.FileName ?? "Unknown");

                        // Skip already-validated or already-problem models
                        if (model.State == SwModelInfo.ProcessingState.Problem)
                        {
                            ErrorHandler.DebugLog($"[BATCHVAL]   Skipping - already marked as Problem");
                            result.ProblemModels.Add(model);
                            result.TotalValidated++;
                            continue;
                        }

                        ValidateSingleModel(model, result);
                        result.TotalValidated++;
                        ErrorHandler.DebugLog($"[BATCHVAL]   After validation: Good={result.GoodModels.Count}, Problem={result.ProblemModels.Count}");
                    }
                    catch (Exception modelEx)
                    {
                        ErrorHandler.HandleError(proc, $"Validation failed for {model.FileName}", modelEx, ErrorHandler.LogLevel.Error);
                        MarkAsProblem(model, $"Validation exception: {modelEx.Message}", ProblemPartManager.ProblemCategory.Fatal);
                        result.ProblemModels.Add(model);
                        result.TotalValidated++;
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Batch validation failed (outer)", ex, ErrorHandler.LogLevel.Error);
            }
            finally
            {
                sw.Stop();
                result.Elapsed = sw.Elapsed;
                ErrorHandler.PopCallStack();
            }

            return result;
        }

        private void ValidateSingleModel(SwModelInfo model, BatchValidationResult result)
        {
            IModelDoc2 doc = null;
            bool needsClose = false;

            ErrorHandler.DebugLog($"[BATCHVAL] ValidateSingleModel: {model.FileName} State={model.State}");

            try
            {
                // Use already-open doc if available (stored in ModelDoc property)
                doc = model.ModelDoc as IModelDoc2;
                ErrorHandler.DebugLog($"[BATCHVAL]   ModelDoc from property: {(doc != null ? "yes" : "null")}");

                if (doc == null && !string.IsNullOrWhiteSpace(model.FilePath))
                {
                    // Open silently for validation (read-only is faster)
                    ErrorHandler.DebugLog($"[BATCHVAL]   Opening file: {model.FilePath}");
                    doc = _fileOps.OpenSWDocument(
                        model.FilePath,
                        silent: true,
                        readOnly: true,
                        configurationName: model.Configuration ?? string.Empty);
                    needsClose = true;
                    ErrorHandler.DebugLog($"[BATCHVAL]   Opened: {(doc != null ? "yes" : "FAILED")}");
                }

                if (doc == null)
                {
                    ErrorHandler.DebugLog($"[BATCHVAL]   PROBLEM: Failed to open file");
                    MarkAsProblem(model, "Failed to open file", ProblemPartManager.ProblemCategory.FileAccess);
                    result.ProblemModels.Add(model);
                    return;
                }

                // Run validation
                var vr = _validator.Validate(model, doc);
                ErrorHandler.DebugLog($"[BATCHVAL]   Validation result: Success={vr.Success}, Summary={vr.Summary}");

                if (vr.Success)
                {
                    model.MarkValidated(true);
                    result.GoodModels.Add(model);
                    ErrorHandler.DebugLog($"[BATCHVAL]   Added to GoodModels, count now: {result.GoodModels.Count}");
                }
                else
                {
                    ErrorHandler.DebugLog($"[BATCHVAL]   PROBLEM: {vr.Summary}");
                    // MarkAsProblem calls model.MarkValidated internally
                    MarkAsProblem(model, vr.Summary, MapReasonToCategory(vr.Summary));
                    result.ProblemModels.Add(model);
                    ErrorHandler.DebugLog($"[BATCHVAL]   Added to ProblemModels, count now: {result.ProblemModels.Count}");
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[BATCHVAL]   EXCEPTION: {ex.Message}");
                MarkAsProblem(model, ex.Message, ProblemPartManager.ProblemCategory.Fatal);
                result.ProblemModels.Add(model);
            }
            finally
            {
                ErrorHandler.DebugLog($"[BATCHVAL]   Finally block, needsClose={needsClose}, doc={(doc != null ? "not null" : "null")}");
                if (needsClose && doc != null)
                {
                    try
                    {
                        var title = doc.GetTitle();
                        ErrorHandler.DebugLog($"[BATCHVAL]   Closing doc: {title}");
                        _swApp.CloseDoc(title);
                        ErrorHandler.DebugLog($"[BATCHVAL]   Doc closed successfully");
                    }
                    catch (Exception closeEx)
                    {
                        ErrorHandler.DebugLog($"[BATCHVAL]   Close failed: {closeEx.Message}");
                    }
                }
            }
        }

        private void MarkAsProblem(SwModelInfo model, string reason, ProblemPartManager.ProblemCategory category)
        {
            model.MarkValidated(false, reason);

            ProblemPartManager.Instance.AddProblemPart(
                model.FilePath ?? string.Empty,
                model.Configuration ?? string.Empty,
                model.ComponentName ?? string.Empty,
                reason,
                category);
        }

        private static ProblemPartManager.ProblemCategory MapReasonToCategory(string reason)
        {
            if (string.IsNullOrWhiteSpace(reason))
                return ProblemPartManager.ProblemCategory.ProcessingError;

            var lower = reason.ToLowerInvariant();

            if (lower.Contains("no solid bod") || lower.Contains("no bodies"))
                return ProblemPartManager.ProblemCategory.GeometryValidation;
            if (lower.Contains("multi-body") || lower.Contains("multiple bod"))
                return ProblemPartManager.ProblemCategory.GeometryValidation;
            if (lower.Contains("material"))
                return ProblemPartManager.ProblemCategory.MaterialMissing;
            if (lower.Contains("thickness"))
                return ProblemPartManager.ProblemCategory.ThicknessExtraction;
            if (lower.Contains("file") || lower.Contains("open") || lower.Contains("access"))
                return ProblemPartManager.ProblemCategory.FileAccess;
            if (lower.Contains("suppress"))
                return ProblemPartManager.ProblemCategory.Suppressed;
            if (lower.Contains("lightweight"))
                return ProblemPartManager.ProblemCategory.Lightweight;
            if (lower.Contains("import"))
                return ProblemPartManager.ProblemCategory.Imported;

            return ProblemPartManager.ProblemCategory.ProcessingError;
        }

        /// <summary>
        /// Re-validates specific models (after user fixes).
        /// Returns models that now pass validation.
        /// </summary>
        public List<SwModelInfo> RevalidateModels(IList<SwModelInfo> models)
        {
            var nowGood = new List<SwModelInfo>();

            foreach (var model in models)
            {
                if (model == null) continue;

                // Reset state to allow re-validation
                model.ResetState();

                var tempResult = new BatchValidationResult();
                ValidateSingleModel(model, tempResult);

                if (tempResult.GoodModels.Count > 0)
                {
                    nowGood.Add(model);
                    // Remove from problem manager
                    var problemItem = FindProblemItem(model);
                    if (problemItem != null)
                    {
                        ProblemPartManager.Instance.RemoveResolvedPart(problemItem);
                    }
                }
            }

            return nowGood;
        }

        private ProblemPartManager.ProblemItem FindProblemItem(SwModelInfo model)
        {
            foreach (var item in ProblemPartManager.Instance.GetProblemParts())
            {
                if (string.Equals(item.FilePath, model.FilePath, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(item.Configuration ?? "", model.Configuration ?? "", StringComparison.OrdinalIgnoreCase))
                {
                    return item;
                }
            }
            return null;
        }
    }
}
