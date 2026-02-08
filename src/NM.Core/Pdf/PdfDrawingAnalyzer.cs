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
        /// Merges AI-extracted data into the DrawingData, preferring higher-confidence values.
        /// </summary>
        private async Task EnhanceWithVisionAsync(DrawingData result, string pdfPath, VisionContext context)
        {
            if (!_visionService.IsAvailable)
                return;

            try
            {
                // For now, we need the caller to provide page images (PDF→PNG rendering).
                // PDFtoImage integration happens in Phase 2.5 when we add the NuGet package.
                // This method is the integration point — the rendering step plugs in here.

                // Placeholder: read the PDF file as bytes for services that accept PDF directly
                byte[] pdfBytes = File.ReadAllBytes(pdfPath);

                // In a full implementation, we would:
                // 1. Render page 1 to PNG at configured DPI
                // 2. Optionally crop to title block region for Tier 2
                // 3. Send to AI service
                // 4. Merge results

                // For now, we'll send the raw bytes and let the service handle it
                // (ClaudeVisionService expects PNG, so this is a no-op until rendering is wired up)

                // The key insight: the architecture is ready. When PDF→PNG rendering is added,
                // it slots in right here with zero changes to the rest of the pipeline.

                System.Diagnostics.Debug.WriteLine(
                    $"[PDF] AI Vision ready. Service: {_visionService.GetType().Name}, " +
                    $"Available: {_visionService.IsAvailable}, " +
                    $"Session cost: ${_visionService.SessionCost:F3}");
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

            // Fill gaps: only overwrite empty fields with AI data
            if (string.IsNullOrEmpty(target.PartNumber) && vision.PartNumber?.HasValue == true)
                target.PartNumber = vision.PartNumber.Value;
            if (string.IsNullOrEmpty(target.Description) && vision.Description?.HasValue == true)
                target.Description = vision.Description.Value;
            if (string.IsNullOrEmpty(target.Revision) && vision.Revision?.HasValue == true)
                target.Revision = vision.Revision.Value;
            if (string.IsNullOrEmpty(target.Material) && vision.Material?.HasValue == true)
                target.Material = vision.Material.Value;
            if (string.IsNullOrEmpty(target.Finish) && vision.Finish?.HasValue == true)
                target.Finish = vision.Finish.Value;
            if (string.IsNullOrEmpty(target.DrawnBy) && vision.DrawnBy?.HasValue == true)
                target.DrawnBy = vision.DrawnBy.Value;
            if (string.IsNullOrEmpty(target.Scale) && vision.Scale?.HasValue == true)
                target.Scale = vision.Scale.Value;
            if (string.IsNullOrEmpty(target.SheetInfo) && vision.Sheet?.HasValue == true)
                target.SheetInfo = vision.Sheet.Value;

            // Merge manufacturing notes from AI (add new ones not already present)
            if (vision.ManufacturingNotes.Count > 0)
            {
                var existingTexts = new HashSet<string>(
                    target.Notes.Select(n => n.Text),
                    StringComparer.OrdinalIgnoreCase);

                foreach (var aiNote in vision.ManufacturingNotes)
                {
                    if (!existingTexts.Contains(aiNote.Text))
                    {
                        NoteCategory category;
                        if (!Enum.TryParse(aiNote.Category, true, out category))
                            category = NoteCategory.General;

                        target.Notes.Add(new DrawingNote(aiNote.Text, category, aiNote.Confidence));
                    }
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
        /// Tries to find a companion PDF for a SolidWorks part file.
        /// Searches for: same name with .pdf extension, or in a "Drawings" subfolder.
        /// </summary>
        public static string FindCompanionPdf(string partFilePath)
        {
            if (string.IsNullOrEmpty(partFilePath)) return null;

            string dir = Path.GetDirectoryName(partFilePath);
            string baseName = Path.GetFileNameWithoutExtension(partFilePath);
            if (dir == null) return null;

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
