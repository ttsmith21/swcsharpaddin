using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NM.Core.AI;
using NM.Core.AI.Models;
using NM.Core.Pdf.Models;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Orchestrates PDF drawing analysis: text extraction, title block parsing,
    /// note extraction, routing hint generation, and optional AI vision analysis.
    /// Supports tiered analysis:
    ///   Tier 1 (TextOnly): PdfPig extraction + regex parsing — free, offline
    ///   Tier 2 (TitleBlockOnly): AI analyzes title block region — ~$0.001/page
    ///   Tier 3 (FullPage): AI analyzes entire drawing — ~$0.004/page
    /// </summary>
    public sealed class PdfDrawingAnalyzer
    {
        private readonly PdfTextExtractor _textExtractor;
        private readonly TitleBlockParser _titleBlockParser;
        private readonly DrawingNoteExtractor _noteExtractor;
        private readonly IDrawingVisionService _visionService;

        /// <summary>
        /// Creates an analyzer with offline-only analysis (Phase 1).
        /// </summary>
        public PdfDrawingAnalyzer()
            : this(new OfflineVisionService())
        {
        }

        /// <summary>
        /// Creates an analyzer with the specified AI vision service (Phase 2).
        /// </summary>
        public PdfDrawingAnalyzer(IDrawingVisionService visionService)
        {
            _textExtractor = new PdfTextExtractor();
            _titleBlockParser = new TitleBlockParser();
            _noteExtractor = new DrawingNoteExtractor();
            _visionService = visionService ?? new OfflineVisionService();
        }

        /// <summary>
        /// The vision service in use (for checking availability and cost).
        /// </summary>
        public IDrawingVisionService VisionService => _visionService;

        /// <summary>
        /// Analyzes a PDF drawing using text extraction only (Phase 1, free).
        /// </summary>
        public DrawingData Analyze(string pdfPath)
        {
            return AnalyzeCore(pdfPath, useVision: false, context: null);
        }

        /// <summary>
        /// Analyzes a PDF drawing with AI vision enhancement (Phase 2).
        /// Falls back to text-only if AI is unavailable.
        /// </summary>
        public DrawingData AnalyzeWithVision(string pdfPath, VisionContext context = null)
        {
            return AnalyzeCore(pdfPath, useVision: true, context: context);
        }

        /// <summary>
        /// Async version for UI contexts. Analyzes with AI vision if available.
        /// </summary>
        public async Task<DrawingData> AnalyzeWithVisionAsync(string pdfPath, VisionContext context = null)
        {
            // Phase 1: text extraction (synchronous, fast)
            var result = AnalyzeCore(pdfPath, useVision: false, context: null);

            // Phase 2: AI vision enhancement (async, costs money)
            if (_visionService.IsAvailable && result.PageCount > 0)
            {
                await EnhanceWithVisionAsync(result, pdfPath, context);
            }

            return result;
        }

        private DrawingData AnalyzeCore(string pdfPath, bool useVision, VisionContext context)
        {
            if (string.IsNullOrEmpty(pdfPath) || !File.Exists(pdfPath))
            {
                return new DrawingData
                {
                    SourcePdfPath = pdfPath,
                    Method = AnalysisMethod.TextOnly,
                    OverallConfidence = 0
                };
            }

            var result = new DrawingData
            {
                SourcePdfPath = pdfPath,
                Method = AnalysisMethod.TextOnly
            };

            // Step 1: Extract text from all pages
            var pages = _textExtractor.ExtractText(pdfPath);
            result.PageCount = pages.Count;

            bool hasText = pages.Count > 0 && pages.Any(p => p.HasText);

            if (hasText)
            {
                // Concatenate all text for raw storage
                result.RawText = string.Join("\n--- PAGE BREAK ---\n",
                    pages.Where(p => p.HasText).Select(p => p.FullText));

                // Step 2: Parse title block (typically on first page)
                var titleBlock = _titleBlockParser.ParseFromPage(pages[0], _textExtractor);
                ApplyTitleBlock(result, titleBlock);

                // If first page had weak results, try other pages too
                if (titleBlock.OverallConfidence < 0.5 && pages.Count > 1)
                {
                    for (int i = 1; i < pages.Count; i++)
                    {
                        var altBlock = _titleBlockParser.ParseFromPage(pages[i], _textExtractor);
                        if (altBlock.OverallConfidence > titleBlock.OverallConfidence)
                        {
                            ApplyTitleBlock(result, altBlock);
                            titleBlock = altBlock;
                        }
                    }
                }

                // Step 3: Extract manufacturing notes from all pages
                foreach (var page in pages.Where(p => p.HasText))
                {
                    var notes = _noteExtractor.ExtractNotes(page.FullText, page.PageNumber);
                    result.Notes.AddRange(notes);
                }

                // Step 4: Generate routing hints from notes
                var routingHints = _noteExtractor.GenerateRoutingHints(result.Notes);
                result.RoutingHints.AddRange(routingHints);

                // Step 5: Calculate overall confidence
                result.OverallConfidence = CalculateOverallConfidence(result, titleBlock);

                // Step 5b: Check coverage density for suspicious extractions
                bool hasTitleBlock = !string.IsNullOrEmpty(result.PartNumber)
                    || !string.IsNullOrEmpty(result.Material)
                    || !string.IsNullOrEmpty(result.Description);
                var density = ConfidenceCalibrator.CheckCoverageDensity(
                    result.PageCount, result.Notes.Count, result.GdtCallouts.Count,
                    !string.IsNullOrEmpty(result.ToleranceGeneral), hasTitleBlock);
                if (density.Suspicious)
                {
                    result.CoverageSuspicious = true;
                    result.Warnings.Add(density.Reason);
                }
            }

            // Step 6: AI Vision enhancement (synchronous path)
            if (useVision && _visionService.IsAvailable)
            {
                bool textExtractionWeak = !hasText || result.OverallConfidence < 0.6;

                // For scanned PDFs or low-confidence text extraction, use AI
                if (textExtractionWeak)
                {
                    try
                    {
                        // The synchronous path — for batch processing or non-UI contexts
                        // In production, prefer AnalyzeWithVisionAsync for responsive UI
                        var visionTask = EnhanceWithVisionAsync(result, pdfPath, context);
                        visionTask.Wait();
                    }
                    catch (AggregateException ae)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[PDF] AI Vision failed: {ae.InnerException?.Message ?? ae.Message}");
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Enhances existing analysis results with AI vision data.
        /// For Gemini: sends PDF bytes directly (native PDF support).
        /// For Claude: renders PDF to PNG first, falls back if no renderer available.
        /// </summary>
        private async Task EnhanceWithVisionAsync(DrawingData result, string pdfPath, VisionContext context)
        {
            if (!_visionService.IsAvailable)
                return;

            try
            {
                byte[] inputBytes;
                bool isGemini = _visionService is GeminiVisionService;

                if (isGemini)
                {
                    // Gemini accepts PDF natively — no rendering needed
                    inputBytes = PdfRenderer.GetPdfBytes(pdfPath);
                    if (inputBytes == null || inputBytes.Length == 0)
                    {
                        System.Diagnostics.Debug.WriteLine("[PDF] Failed to read PDF bytes");
                        return;
                    }
                }
                else
                {
                    // Claude/other services need PNG — try rendering
                    inputBytes = PdfRenderer.RenderPageToPng(pdfPath, pageIndex: 0, dpi: 200);
                    if (inputBytes == null || inputBytes.Length == 0)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            "[PDF] No PNG renderer available. Install Ghostscript for Claude Vision, " +
                            "or switch to Gemini provider which accepts PDF natively.");
                        return;
                    }
                }

                VisionAnalysisResult visionResult;

                // Choose analysis tier based on text extraction confidence
                bool titleBlockOnly = result.OverallConfidence >= 0.6;
                if (titleBlockOnly && !isGemini)
                {
                    // Tier 2: title block region only (cheaper, Claude only)
                    byte[] titleBlockImage = PdfRenderer.RenderTitleBlockRegion(pdfPath, pageIndex: 0, dpi: 300);
                    byte[] imageToAnalyze = (titleBlockImage != null && titleBlockImage.Length > 0)
                        ? titleBlockImage : inputBytes;
                    visionResult = await _visionService.AnalyzeTitleBlockAsync(imageToAnalyze);
                }
                else
                {
                    // Tier 3: full page/document analysis
                    visionResult = await _visionService.AnalyzeDrawingPageAsync(inputBytes, context);
                }

                if (visionResult != null && visionResult.Success)
                {
                    MergeVisionResults(result, visionResult);

                    // Re-check coverage density after AI merge (AI may have added notes)
                    bool hasTitleBlock = !string.IsNullOrEmpty(result.PartNumber)
                        || !string.IsNullOrEmpty(result.Material)
                        || !string.IsNullOrEmpty(result.Description);
                    var density = ConfidenceCalibrator.CheckCoverageDensity(
                        result.PageCount, result.Notes.Count, result.GdtCallouts.Count,
                        !string.IsNullOrEmpty(result.ToleranceGeneral), hasTitleBlock);
                    if (density.Suspicious && !result.CoverageSuspicious)
                    {
                        result.CoverageSuspicious = true;
                        result.Warnings.Add(density.Reason);
                    }
                    else if (!density.Suspicious && result.CoverageSuspicious)
                    {
                        // AI resolved the coverage issue — clear the flag
                        result.CoverageSuspicious = false;
                    }
                }

                System.Diagnostics.Debug.WriteLine(
                    $"[PDF] AI Vision complete. Service: {_visionService.GetType().Name}, " +
                    $"Success: {visionResult?.Success}, " +
                    $"Cost: ${visionResult?.CostUsd:F3}, " +
                    $"Session total: ${_visionService.SessionCost:F3}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PDF] AI Vision enhancement failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Merges AI vision results into the DrawingData, filling gaps and improving confidence.
        /// Called after AI analysis completes.
        /// </summary>
        public void MergeVisionResults(DrawingData target, VisionAnalysisResult vision)
        {
            if (target == null || vision == null || !vision.Success)
                return;

            // Cross-validated merge: compare text vs vision for each title block field
            MergeField(target, "PartNumber", target.PartNumber, vision.PartNumber,
                v => target.PartNumber = v);
            MergeField(target, "Description", target.Description, vision.Description,
                v => target.Description = v);
            MergeField(target, "Revision", target.Revision, vision.Revision,
                v => target.Revision = v);
            MergeField(target, "Material", target.Material, vision.Material,
                v => target.Material = v);
            MergeField(target, "Finish", target.Finish, vision.Finish,
                v => target.Finish = v);
            MergeField(target, "DrawnBy", target.DrawnBy, vision.DrawnBy,
                v => target.DrawnBy = v);
            MergeField(target, "Scale", target.Scale, vision.Scale,
                v => target.Scale = v);
            MergeField(target, "SheetInfo", target.SheetInfo, vision.Sheet,
                v => target.SheetInfo = v);

            // Merge manufacturing notes from AI (add new ones not already present)
            if (vision.ManufacturingNotes.Count > 0)
            {
                var existingTexts = new HashSet<string>(
                    target.Notes.Select(n => n.Text),
                    StringComparer.OrdinalIgnoreCase);

                foreach (var aiNote in vision.ManufacturingNotes)
                {
                    // Check if AI note matches an existing text note (case-insensitive)
                    bool matchesExisting = existingTexts.Contains(aiNote.Text);
                    if (!matchesExisting)
                    {
                        NoteCategory category;
                        if (!Enum.TryParse(aiNote.Category, true, out category))
                            category = NoteCategory.General;

                        target.Notes.Add(new DrawingNote(aiNote.Text, category, aiNote.Confidence));
                    }
                    // When AI confirms a text note, the agreement implicitly boosts
                    // confidence (cross-validation at note level is informational only)
                }

                // Regenerate routing hints with the merged notes
                var hints = _noteExtractor.GenerateRoutingHints(target.Notes);
                target.RoutingHints.Clear();
                target.RoutingHints.AddRange(hints);
            }

            // Update method and confidence
            target.Method = string.IsNullOrEmpty(target.RawText)
                ? AnalysisMethod.VisionAI
                : AnalysisMethod.Hybrid;

            // Boost confidence when AI confirms text extraction
            if (vision.OverallConfidence > target.OverallConfidence)
                target.OverallConfidence = vision.OverallConfidence;
        }

        /// <summary>
        /// Cross-validates a single title block field between text and vision sources.
        /// When both agree, confidence is boosted. When they disagree, a warning is added.
        /// When only one source has a value, it's used with reduced confidence.
        /// </summary>
        private static void MergeField(
            DrawingData target, string fieldName,
            string textValue, FieldResult visionField,
            Action<string> setter)
        {
            bool textHas = !string.IsNullOrEmpty(textValue);
            bool visionHas = visionField?.HasValue == true;

            if (textHas && visionHas)
            {
                bool agree = string.Equals(textValue.Trim(), visionField.Value.Trim(),
                    StringComparison.OrdinalIgnoreCase);
                if (!agree)
                {
                    // Conflict — keep text value (text is more reliable for structured fields)
                    target.Warnings.Add(
                        $"{fieldName} conflict: text='{textValue}', vision='{visionField.Value}'");
                }
                // Both agree or conflict: keep text value, cross-validation adjusts overall confidence
            }
            else if (!textHas && visionHas)
            {
                // Gap fill from vision
                setter(visionField.Value);
            }
            // If only text has value, keep it as-is
        }

        /// <summary>
        /// Tries to find a companion PDF for a SolidWorks part file.
        /// Searches for: same name with .pdf extension, or in a "Drawings" subfolder.
        /// </summary>
        public static string FindCompanionPdf(string partFilePath)
        {
            if (string.IsNullOrEmpty(partFilePath)) return null;

            string dir = Path.GetDirectoryName(partFilePath);
            string baseName = Path.GetFileNameWithoutExtension(partFilePath);
            if (dir == null || !Directory.Exists(dir)) return null;

            // Same directory, same name
            string sameName = Path.Combine(dir, baseName + ".pdf");
            if (File.Exists(sameName)) return sameName;

            // Same directory, case-insensitive search
            var pdfFiles = Directory.GetFiles(dir, "*.pdf", SearchOption.TopDirectoryOnly);
            var match = pdfFiles.FirstOrDefault(f =>
                string.Equals(Path.GetFileNameWithoutExtension(f), baseName, StringComparison.OrdinalIgnoreCase));
            if (match != null) return match;

            // Check "Drawings" subfolder
            string drawingsDir = Path.Combine(dir, "Drawings");
            if (Directory.Exists(drawingsDir))
            {
                sameName = Path.Combine(drawingsDir, baseName + ".pdf");
                if (File.Exists(sameName)) return sameName;

                pdfFiles = Directory.GetFiles(drawingsDir, "*.pdf", SearchOption.TopDirectoryOnly);
                match = pdfFiles.FirstOrDefault(f =>
                    string.Equals(Path.GetFileNameWithoutExtension(f), baseName, StringComparison.OrdinalIgnoreCase));
                if (match != null) return match;
            }

            // Check "PDF" subfolder
            string pdfDir = Path.Combine(dir, "PDF");
            if (Directory.Exists(pdfDir))
            {
                sameName = Path.Combine(pdfDir, baseName + ".pdf");
                if (File.Exists(sameName)) return sameName;

                pdfFiles = Directory.GetFiles(pdfDir, "*.pdf", SearchOption.TopDirectoryOnly);
                match = pdfFiles.FirstOrDefault(f =>
                    string.Equals(Path.GetFileNameWithoutExtension(f), baseName, StringComparison.OrdinalIgnoreCase));
                if (match != null) return match;
            }

            // Check parent directory
            string parentDir = Directory.GetParent(dir)?.FullName;
            if (parentDir != null)
            {
                sameName = Path.Combine(parentDir, baseName + ".pdf");
                if (File.Exists(sameName)) return sameName;
            }

            return null;
        }

        /// <summary>
        /// Quick check: does this PDF have extractable text?
        /// Returns false for scanned/image-only PDFs.
        /// </summary>
        public bool HasExtractableText(string pdfPath)
        {
            var pages = _textExtractor.ExtractText(pdfPath);
            return pages.Any(p => p.HasText && p.FullText.Length > 20);
        }

        private void ApplyTitleBlock(DrawingData result, TitleBlockInfo titleBlock)
        {
            if (!string.IsNullOrEmpty(titleBlock.PartNumber))
                result.PartNumber = titleBlock.PartNumber;
            if (!string.IsNullOrEmpty(titleBlock.Description))
                result.Description = titleBlock.Description;
            if (!string.IsNullOrEmpty(titleBlock.Revision))
                result.Revision = titleBlock.Revision;
            if (!string.IsNullOrEmpty(titleBlock.Material))
                result.Material = titleBlock.Material;
            if (!string.IsNullOrEmpty(titleBlock.Finish))
                result.Finish = titleBlock.Finish;
            if (!string.IsNullOrEmpty(titleBlock.DrawnBy))
                result.DrawnBy = titleBlock.DrawnBy;
            if (!string.IsNullOrEmpty(titleBlock.CheckedBy))
                result.CheckedBy = titleBlock.CheckedBy;
            if (titleBlock.Date.HasValue)
                result.DrawingDate = titleBlock.Date;
            if (!string.IsNullOrEmpty(titleBlock.Scale))
                result.Scale = titleBlock.Scale;
            if (!string.IsNullOrEmpty(titleBlock.Sheet))
                result.SheetInfo = titleBlock.Sheet;
            if (!string.IsNullOrEmpty(titleBlock.ToleranceGeneral))
                result.ToleranceGeneral = titleBlock.ToleranceGeneral;
        }

        private double CalculateOverallConfidence(DrawingData result, TitleBlockInfo titleBlock)
        {
            double score = 0;
            int factors = 0;

            // Title block confidence
            if (titleBlock.OverallConfidence > 0)
            {
                score += titleBlock.OverallConfidence;
                factors++;
            }

            // Bonus for key fields being populated
            if (!string.IsNullOrEmpty(result.PartNumber)) { score += 0.8; factors++; }
            if (!string.IsNullOrEmpty(result.Material)) { score += 0.8; factors++; }
            if (!string.IsNullOrEmpty(result.Description)) { score += 0.7; factors++; }
            if (!string.IsNullOrEmpty(result.Revision)) { score += 0.9; factors++; }

            // Note extraction adds confidence
            if (result.Notes.Count > 0)
            {
                score += 0.7;
                factors++;
            }

            return factors > 0 ? score / factors : 0;
        }
    }
}
