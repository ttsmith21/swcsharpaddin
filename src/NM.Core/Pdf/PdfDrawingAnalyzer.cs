using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NM.Core.Pdf.Models;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Orchestrates PDF drawing analysis: text extraction, title block parsing,
    /// note extraction, and routing hint generation.
    /// This is the main entry point for Phase 1 (offline analysis).
    /// </summary>
    public sealed class PdfDrawingAnalyzer
    {
        private readonly PdfTextExtractor _textExtractor;
        private readonly TitleBlockParser _titleBlockParser;
        private readonly DrawingNoteExtractor _noteExtractor;

        public PdfDrawingAnalyzer()
        {
            _textExtractor = new PdfTextExtractor();
            _titleBlockParser = new TitleBlockParser();
            _noteExtractor = new DrawingNoteExtractor();
        }

        /// <summary>
        /// Analyzes a PDF drawing and returns all extracted data.
        /// </summary>
        public DrawingData Analyze(string pdfPath)
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

            if (pages.Count == 0 || pages.All(p => !p.HasText))
            {
                // No text extracted — this is likely a scanned PDF
                // Phase 2 (AI Vision) would handle this
                result.OverallConfidence = 0;
                return result;
            }

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

            return result;
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
