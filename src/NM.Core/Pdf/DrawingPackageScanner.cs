using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using NM.Core.Pdf.Models;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Scans a folder of PDFs (or a single multi-page PDF) and builds a DrawingPackageIndex
    /// mapping part numbers to their drawing pages. Handles the reality of customer drawing
    /// packages: one PDF per part, one PDF for the whole assembly, or anything in between.
    /// </summary>
    public sealed class DrawingPackageScanner
    {
        private readonly PdfTextExtractor _textExtractor;
        private readonly TitleBlockParser _titleBlockParser;
        private readonly DrawingNoteExtractor _noteExtractor;
        private readonly SpecRecognizer _specRecognizer;
        private readonly ToleranceAnalyzer _toleranceAnalyzer;

        public DrawingPackageScanner()
        {
            _textExtractor = new PdfTextExtractor();
            _titleBlockParser = new TitleBlockParser();
            _noteExtractor = new DrawingNoteExtractor();
            _specRecognizer = new SpecRecognizer();
            _toleranceAnalyzer = new ToleranceAnalyzer();
        }

        /// <summary>
        /// Scans all PDFs in a folder and builds an index of drawing pages by part number.
        /// </summary>
        public DrawingPackageIndex ScanFolder(string folderPath)
        {
            var index = new DrawingPackageIndex();

            if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
                return index;

            var pdfFiles = Directory.GetFiles(folderPath, "*.pdf", SearchOption.TopDirectoryOnly)
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var pdfFile in pdfFiles)
            {
                ScanSinglePdf(pdfFile, index);
            }

            return index;
        }

        /// <summary>
        /// Scans a single PDF (possibly multi-page) and adds pages to the index.
        /// </summary>
        public void ScanSinglePdf(string pdfPath, DrawingPackageIndex index)
        {
            if (string.IsNullOrEmpty(pdfPath) || !File.Exists(pdfPath))
                return;

            index.ScannedFiles.Add(pdfPath);

            List<PageText> pages;
            try
            {
                pages = _textExtractor.ExtractText(pdfPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PKG] Failed to extract text from {pdfPath}: {ex.Message}");
                return;
            }

            index.TotalPages += pages.Count;

            foreach (var page in pages)
            {
                var pageInfo = AnalyzePage(page, pdfPath);

                if (!string.IsNullOrWhiteSpace(pageInfo.PartNumber))
                {
                    if (!index.PagesByPartNumber.TryGetValue(pageInfo.PartNumber, out var list))
                    {
                        list = new List<DrawingPageInfo>();
                        index.PagesByPartNumber[pageInfo.PartNumber] = list;
                    }
                    list.Add(pageInfo);
                }
                else
                {
                    index.UnmatchedPages.Add(pageInfo);
                }

                // Aggregate BOM entries
                if (pageInfo.BomEntries.Count > 0)
                {
                    index.AllBomEntries.AddRange(pageInfo.BomEntries);
                }
            }
        }

        /// <summary>
        /// Scans a list of specific PDF files and builds an index.
        /// Useful when user selects specific files rather than a whole folder.
        /// </summary>
        public DrawingPackageIndex ScanFiles(IEnumerable<string> pdfPaths)
        {
            var index = new DrawingPackageIndex();

            if (pdfPaths == null)
                return index;

            foreach (var pdfPath in pdfPaths)
            {
                ScanSinglePdf(pdfPath, index);
            }

            return index;
        }

        /// <summary>
        /// Analyzes a single page, extracting title block data, notes, and BOM info.
        /// </summary>
        public DrawingPageInfo AnalyzePage(PageText page, string pdfPath)
        {
            var pageInfo = new DrawingPageInfo
            {
                PdfPath = pdfPath,
                PageNumber = page.PageNumber,
                HasText = page.HasText
            };

            if (!page.HasText)
                return pageInfo;

            // Parse title block
            var titleBlock = _titleBlockParser.ParseFromPage(page, _textExtractor);
            pageInfo.PartNumber = titleBlock.PartNumber;
            pageInfo.Description = titleBlock.Description;
            pageInfo.Revision = titleBlock.Revision;
            pageInfo.Material = titleBlock.Material;
            pageInfo.SheetInfo = titleBlock.Sheet;
            pageInfo.Confidence = titleBlock.OverallConfidence;

            // Extract notes
            var notes = _noteExtractor.ExtractNotes(page.FullText, page.PageNumber);
            foreach (var note in notes)
            {
                pageInfo.Notes.Add(note);
            }

            // Recognize formal specifications
            var specMatches = _specRecognizer.Recognize(page.FullText);
            if (specMatches.Count > 0)
            {
                pageInfo.RecognizedSpecs.AddRange(specMatches);

                var specHints = _specRecognizer.ToRoutingHints(specMatches);
                foreach (var hint in specHints)
                {
                    pageInfo.RoutingHints.Add(hint);
                }
            }

            // Generate routing hints from notes
            var hints = _noteExtractor.GenerateRoutingHints(notes);
            foreach (var hint in hints)
            {
                pageInfo.RoutingHints.Add(hint);
            }

            // Tolerance analysis
            var tolResult = _toleranceAnalyzer.Analyze(page.FullText, titleBlock.ToleranceGeneral);
            pageInfo.ToleranceAnalysis = tolResult;
            if (tolResult.CostFlags.Count > 0)
            {
                foreach (var tolHint in _toleranceAnalyzer.ToRoutingHints(tolResult))
                {
                    pageInfo.RoutingHints.Add(tolHint);
                }
            }

            // Detect BOM presence (indicates assembly-level page)
            DetectBom(page, pageInfo);

            // Detect assembly-level drawing
            pageInfo.IsAssemblyLevel = pageInfo.HasBom || IsAssemblyDrawing(page.FullText);

            return pageInfo;
        }

        /// <summary>
        /// Detects BOM table rows in page text using common patterns.
        /// </summary>
        private void DetectBom(PageText page, DrawingPageInfo pageInfo)
        {
            string text = page.FullText;
            if (string.IsNullOrWhiteSpace(text))
                return;

            // Check for BOM header keywords
            bool hasBomHeader = Regex.IsMatch(text,
                @"\b(?:BILL\s*OF\s*MATERIAL|BOM|PARTS?\s*LIST|ITEM\s+(?:NO|#|NUMBER))\b",
                RegexOptions.IgnoreCase);

            if (!hasBomHeader)
                return;

            pageInfo.HasBom = true;

            // Try to extract BOM rows: ITEM# PART# DESCRIPTION QTY
            var bomRowPattern = new Regex(
                @"^\s*(\d{1,3})\s+([A-Z0-9][\w\-\.]+)\s+(.+?)\s+(\d{1,4})\s*$",
                RegexOptions.IgnoreCase | RegexOptions.Multiline);

            var matches = bomRowPattern.Matches(text);
            foreach (Match match in matches)
            {
                var entry = new BomEntry
                {
                    ItemNumber = match.Groups[1].Value.Trim(),
                    PartNumber = match.Groups[2].Value.Trim(),
                    Description = match.Groups[3].Value.Trim(),
                    Quantity = int.TryParse(match.Groups[4].Value, out int qty) ? qty : 1,
                    Confidence = 0.70
                };
                pageInfo.BomEntries.Add(entry);
            }
        }

        /// <summary>
        /// Heuristic: does this page text suggest an assembly-level drawing?
        /// </summary>
        private static bool IsAssemblyDrawing(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            return Regex.IsMatch(text,
                @"\b(?:ASSEMBLY|ASSY|WELDMENT|WELDED\s+ASSY|SUB[\-\s]?ASSY)\b",
                RegexOptions.IgnoreCase);
        }
    }
}
