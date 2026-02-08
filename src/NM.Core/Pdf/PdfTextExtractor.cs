using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Extracts text from PDF files using PdfPig.
    /// Works offline with digitally-created PDFs (not scanned images).
    /// </summary>
    public sealed class PdfTextExtractor
    {
        /// <summary>
        /// Extracts all text from a PDF file, page by page.
        /// Returns empty list if the file cannot be read or contains no text.
        /// </summary>
        public List<PageText> ExtractText(string pdfPath)
        {
            if (string.IsNullOrEmpty(pdfPath) || !File.Exists(pdfPath))
                return new List<PageText>();

            try
            {
                var pages = new List<PageText>();
                using (var document = PdfDocument.Open(pdfPath))
                {
                    foreach (var page in document.GetPages())
                    {
                        var text = ContentOrderTextExtractor.GetText(page);
                        var words = page.GetWords().ToList();

                        pages.Add(new PageText
                        {
                            PageNumber = page.Number,
                            FullText = text ?? string.Empty,
                            Words = words.Select(w => new WordInfo
                            {
                                Text = w.Text,
                                Left = w.BoundingBox.Left,
                                Bottom = w.BoundingBox.Bottom,
                                Right = w.BoundingBox.Right,
                                Top = w.BoundingBox.Top,
                                FontSize = w.Letters.FirstOrDefault()?.PointSize ?? 0
                            }).ToList(),
                            Width = page.Width,
                            Height = page.Height
                        });
                    }
                }
                return pages;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PDF] Failed to extract text from {pdfPath}: {ex.Message}");
                return new List<PageText>();
            }
        }

        /// <summary>
        /// Extracts text from the bottom-right quadrant of a page (typical title block location).
        /// </summary>
        public string ExtractTitleBlockRegion(PageText page)
        {
            if (page == null || page.Words == null || page.Words.Count == 0)
                return string.Empty;

            // Title blocks are typically in the bottom-right quadrant
            double midX = page.Width * 0.45;  // Right ~55% of page
            double midY = page.Height * 0.35; // Bottom ~35% of page

            var titleBlockWords = page.Words
                .Where(w => w.Left >= midX && w.Bottom <= midY)
                .OrderByDescending(w => w.Top)
                .ThenBy(w => w.Left)
                .Select(w => w.Text);

            return string.Join(" ", titleBlockWords);
        }

        /// <summary>
        /// Extracts text from the upper portion of a page (typical notes area).
        /// </summary>
        public string ExtractNotesRegion(PageText page)
        {
            if (page == null || page.Words == null || page.Words.Count == 0)
                return string.Empty;

            // Notes are often in the upper-left or along the left margin
            double notesX = page.Width * 0.50;  // Left ~50% of page
            double notesY = page.Height * 0.40; // Top ~60% of page

            var noteWords = page.Words
                .Where(w => w.Left <= notesX && w.Top >= notesY)
                .OrderByDescending(w => w.Top)
                .ThenBy(w => w.Left)
                .Select(w => w.Text);

            return string.Join(" ", noteWords);
        }

        /// <summary>
        /// Gets the total page count without extracting text.
        /// </summary>
        public int GetPageCount(string pdfPath)
        {
            if (string.IsNullOrEmpty(pdfPath) || !File.Exists(pdfPath))
                return 0;

            try
            {
                using (var document = PdfDocument.Open(pdfPath))
                {
                    return document.NumberOfPages;
                }
            }
            catch
            {
                return 0;
            }
        }
    }

    /// <summary>
    /// Text content from a single PDF page, with word positions.
    /// </summary>
    public sealed class PageText
    {
        public int PageNumber { get; set; }
        public string FullText { get; set; }
        public List<WordInfo> Words { get; set; } = new List<WordInfo>();
        public double Width { get; set; }
        public double Height { get; set; }

        /// <summary>True if the page has meaningful text content (not just whitespace).</summary>
        public bool HasText => !string.IsNullOrWhiteSpace(FullText);
    }

    /// <summary>
    /// A single word with its bounding box position on the page.
    /// </summary>
    public sealed class WordInfo
    {
        public string Text { get; set; }
        public double Left { get; set; }
        public double Bottom { get; set; }
        public double Right { get; set; }
        public double Top { get; set; }
        public double FontSize { get; set; }
    }
}
