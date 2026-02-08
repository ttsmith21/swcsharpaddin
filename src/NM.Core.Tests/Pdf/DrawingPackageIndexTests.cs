using System.Collections.Generic;
using Xunit;
using NM.Core.Pdf.Models;

namespace NM.Core.Tests.Pdf
{
    public class DrawingPackageIndexTests
    {
        // --- FindPages ---

        [Fact]
        public void FindPages_ExactMatch_ReturnsPages()
        {
            var index = new DrawingPackageIndex();
            var pages = new List<DrawingPageInfo>
            {
                new DrawingPageInfo { PartNumber = "12345-01", PageNumber = 1 }
            };
            index.PagesByPartNumber["12345-01"] = pages;

            var result = index.FindPages("12345-01");
            Assert.Single(result);
        }

        [Fact]
        public void FindPages_CaseInsensitive_Matches()
        {
            var index = new DrawingPackageIndex();
            index.PagesByPartNumber["ABC-123"] = new List<DrawingPageInfo>
            {
                new DrawingPageInfo { PartNumber = "ABC-123", PageNumber = 1 }
            };

            var result = index.FindPages("abc-123");
            Assert.Single(result);
        }

        [Fact]
        public void FindPages_TrimmedMatch_Matches()
        {
            var index = new DrawingPackageIndex();
            index.PagesByPartNumber["12345"] = new List<DrawingPageInfo>
            {
                new DrawingPageInfo { PartNumber = "12345", PageNumber = 1 }
            };

            var result = index.FindPages("  12345  ");
            Assert.Single(result);
        }

        [Fact]
        public void FindPages_PartialMatch_KeyContainsSearch()
        {
            var index = new DrawingPackageIndex();
            index.PagesByPartNumber["12345-01-REV-C"] = new List<DrawingPageInfo>
            {
                new DrawingPageInfo { PartNumber = "12345-01-REV-C", PageNumber = 1 }
            };

            var result = index.FindPages("12345-01");
            Assert.Single(result);
        }

        [Fact]
        public void FindPages_PartialMatch_SearchContainsKey()
        {
            var index = new DrawingPackageIndex();
            index.PagesByPartNumber["12345"] = new List<DrawingPageInfo>
            {
                new DrawingPageInfo { PartNumber = "12345", PageNumber = 1 }
            };

            var result = index.FindPages("12345-01-FULL");
            Assert.Single(result);
        }

        [Fact]
        public void FindPages_NoMatch_ReturnsEmpty()
        {
            var index = new DrawingPackageIndex();
            index.PagesByPartNumber["12345"] = new List<DrawingPageInfo>
            {
                new DrawingPageInfo { PartNumber = "12345", PageNumber = 1 }
            };

            var result = index.FindPages("99999");
            Assert.Empty(result);
        }

        [Fact]
        public void FindPages_NullInput_ReturnsEmpty()
        {
            var index = new DrawingPackageIndex();
            Assert.Empty(index.FindPages(null));
        }

        [Fact]
        public void FindPages_EmptyInput_ReturnsEmpty()
        {
            var index = new DrawingPackageIndex();
            Assert.Empty(index.FindPages(""));
        }

        [Fact]
        public void FindPages_WhitespaceInput_ReturnsEmpty()
        {
            var index = new DrawingPackageIndex();
            Assert.Empty(index.FindPages("   "));
        }

        // --- BuildDrawingData ---

        [Fact]
        public void BuildDrawingData_SinglePage_ReturnsDrawingData()
        {
            var index = new DrawingPackageIndex();
            index.PagesByPartNumber["12345"] = new List<DrawingPageInfo>
            {
                new DrawingPageInfo
                {
                    PartNumber = "12345",
                    Description = "BRACKET",
                    Revision = "A",
                    Material = "A36",
                    PdfPath = "/test.pdf",
                    PageNumber = 1,
                    Confidence = 0.85
                }
            };

            var data = index.BuildDrawingData("12345");

            Assert.NotNull(data);
            Assert.Equal("12345", data.PartNumber);
            Assert.Equal("BRACKET", data.Description);
            Assert.Equal("A", data.Revision);
            Assert.Equal("A36", data.Material);
            Assert.Equal("/test.pdf", data.SourcePdfPath);
            Assert.Equal(1, data.PageCount);
        }

        [Fact]
        public void BuildDrawingData_MultiplePages_MergesNotes()
        {
            var index = new DrawingPackageIndex();
            var page1 = new DrawingPageInfo
            {
                PartNumber = "12345",
                Description = "BRACKET",
                PageNumber = 1,
                Confidence = 0.85,
                PdfPath = "/test.pdf"
            };
            page1.Notes.Add(new DrawingNote { Text = "BREAK ALL EDGES", Category = NoteCategory.Deburr });

            var page2 = new DrawingPageInfo
            {
                PartNumber = "12345",
                PageNumber = 2,
                Confidence = 0.80,
                PdfPath = "/test.pdf"
            };
            page2.Notes.Add(new DrawingNote { Text = "POWDER COAT", Category = NoteCategory.Finish });

            index.PagesByPartNumber["12345"] = new List<DrawingPageInfo> { page1, page2 };

            var data = index.BuildDrawingData("12345");

            Assert.NotNull(data);
            Assert.Equal(2, data.PageCount);
            Assert.Equal(2, data.Notes.Count);
        }

        [Fact]
        public void BuildDrawingData_DeduplicatesNotes()
        {
            var index = new DrawingPackageIndex();
            var page1 = new DrawingPageInfo { PartNumber = "12345", PageNumber = 1, Confidence = 0.85, PdfPath = "/test.pdf" };
            page1.Notes.Add(new DrawingNote { Text = "BREAK ALL EDGES" });

            var page2 = new DrawingPageInfo { PartNumber = "12345", PageNumber = 2, Confidence = 0.80, PdfPath = "/test.pdf" };
            page2.Notes.Add(new DrawingNote { Text = "BREAK ALL EDGES" }); // Duplicate

            index.PagesByPartNumber["12345"] = new List<DrawingPageInfo> { page1, page2 };

            var data = index.BuildDrawingData("12345");

            Assert.Single(data.Notes); // Deduped
        }

        [Fact]
        public void BuildDrawingData_NoMatch_ReturnsNull()
        {
            var index = new DrawingPackageIndex();
            Assert.Null(index.BuildDrawingData("99999"));
        }

        [Fact]
        public void BuildDrawingData_MergesBomEntries()
        {
            var index = new DrawingPackageIndex();
            var page = new DrawingPageInfo
            {
                PartNumber = "ASSY-100",
                PageNumber = 1,
                Confidence = 0.80,
                PdfPath = "/assy.pdf",
                IsAssemblyLevel = true,
                HasBom = true
            };
            page.BomEntries.Add(new BomEntry { PartNumber = "12345-01", Quantity = 2 });
            page.BomEntries.Add(new BomEntry { PartNumber = "12345-02", Quantity = 1 });

            index.PagesByPartNumber["ASSY-100"] = new List<DrawingPageInfo> { page };

            var data = index.BuildDrawingData("ASSY-100");

            Assert.Equal(2, data.BomEntries.Count);
        }

        // --- Computed properties ---

        [Fact]
        public void MatchedPages_ReturnsCorrectCount()
        {
            var index = new DrawingPackageIndex();
            index.PagesByPartNumber["A"] = new List<DrawingPageInfo>
            {
                new DrawingPageInfo { PageNumber = 1 },
                new DrawingPageInfo { PageNumber = 2 }
            };
            index.PagesByPartNumber["B"] = new List<DrawingPageInfo>
            {
                new DrawingPageInfo { PageNumber = 1 }
            };

            Assert.Equal(3, index.MatchedPages);
        }

        [Fact]
        public void UniquePartNumbers_ReturnsCorrectCount()
        {
            var index = new DrawingPackageIndex();
            index.PagesByPartNumber["A"] = new List<DrawingPageInfo>();
            index.PagesByPartNumber["B"] = new List<DrawingPageInfo>();
            index.PagesByPartNumber["C"] = new List<DrawingPageInfo>();

            Assert.Equal(3, index.UniquePartNumbers);
        }

        [Fact]
        public void Summary_FormatsCorrectly()
        {
            var index = new DrawingPackageIndex();
            index.ScannedFiles.Add("/a.pdf");
            index.ScannedFiles.Add("/b.pdf");
            index.TotalPages = 5;
            index.PagesByPartNumber["X"] = new List<DrawingPageInfo>
            {
                new DrawingPageInfo { PageNumber = 1 },
                new DrawingPageInfo { PageNumber = 2 }
            };
            index.UnmatchedPages.Add(new DrawingPageInfo());

            string summary = index.Summary;
            Assert.Contains("2 PDF(s)", summary);
            Assert.Contains("5 pages", summary);
            Assert.Contains("1 part numbers", summary);
            Assert.Contains("1 unmatched", summary);
        }
    }
}
