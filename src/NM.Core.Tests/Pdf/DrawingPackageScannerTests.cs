using System.IO;
using System.Linq;
using Xunit;
using NM.Core.Pdf;
using NM.Core.Pdf.Models;

namespace NM.Core.Tests.Pdf
{
    public class DrawingPackageScannerTests
    {
        private readonly DrawingPackageScanner _scanner = new DrawingPackageScanner();

        // --- ScanFolder ---

        [Fact]
        public void ScanFolder_NullPath_ReturnsEmptyIndex()
        {
            var index = _scanner.ScanFolder(null);
            Assert.NotNull(index);
            Assert.Equal(0, index.TotalPages);
            Assert.Empty(index.ScannedFiles);
        }

        [Fact]
        public void ScanFolder_EmptyPath_ReturnsEmptyIndex()
        {
            var index = _scanner.ScanFolder("");
            Assert.NotNull(index);
            Assert.Equal(0, index.TotalPages);
        }

        [Fact]
        public void ScanFolder_NonexistentPath_ReturnsEmptyIndex()
        {
            var index = _scanner.ScanFolder(Path.Combine(Path.GetTempPath(), "nonexistent_folder_12345"));
            Assert.NotNull(index);
            Assert.Equal(0, index.TotalPages);
        }

        [Fact]
        public void ScanFolder_EmptyFolder_ReturnsEmptyIndex()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), "pkg_test_empty_" + Path.GetRandomFileName());
            Directory.CreateDirectory(tempDir);
            try
            {
                var index = _scanner.ScanFolder(tempDir);
                Assert.NotNull(index);
                Assert.Equal(0, index.TotalPages);
                Assert.Empty(index.ScannedFiles);
            }
            finally
            {
                Directory.Delete(tempDir, true);
            }
        }

        // --- ScanSinglePdf ---

        [Fact]
        public void ScanSinglePdf_NullPath_DoesNotThrow()
        {
            var index = new DrawingPackageIndex();
            _scanner.ScanSinglePdf(null, index);
            Assert.Equal(0, index.TotalPages);
        }

        [Fact]
        public void ScanSinglePdf_NonexistentFile_DoesNotThrow()
        {
            var index = new DrawingPackageIndex();
            _scanner.ScanSinglePdf("/nonexistent/file.pdf", index);
            Assert.Equal(0, index.TotalPages);
        }

        // --- ScanFiles ---

        [Fact]
        public void ScanFiles_NullList_ReturnsEmptyIndex()
        {
            var index = _scanner.ScanFiles(null);
            Assert.NotNull(index);
            Assert.Equal(0, index.TotalPages);
        }

        [Fact]
        public void ScanFiles_EmptyList_ReturnsEmptyIndex()
        {
            var index = _scanner.ScanFiles(new string[0]);
            Assert.NotNull(index);
            Assert.Equal(0, index.TotalPages);
        }

        // --- AnalyzePage ---

        [Fact]
        public void AnalyzePage_EmptyPage_ReturnsPageInfoWithNoPartNumber()
        {
            var page = new PageText
            {
                PageNumber = 1,
                FullText = "",
                Width = 612,
                Height = 792
            };

            var result = _scanner.AnalyzePage(page, "/test.pdf");

            Assert.Equal("/test.pdf", result.PdfPath);
            Assert.Equal(1, result.PageNumber);
            Assert.Null(result.PartNumber);
            Assert.False(result.HasText);
        }

        [Fact]
        public void AnalyzePage_WithPartNumber_ExtractsPartNumber()
        {
            var page = new PageText
            {
                PageNumber = 1,
                FullText = "PART NO: 12345-01\nDESCRIPTION: MOUNTING BRACKET\nREV: C\nMATERIAL: A36 STEEL",
                Width = 612,
                Height = 792
            };

            var result = _scanner.AnalyzePage(page, "/test.pdf");

            Assert.True(result.HasText);
            Assert.Equal("12345-01", result.PartNumber);
        }

        [Fact]
        public void AnalyzePage_WithDescription_ExtractsDescription()
        {
            var page = new PageText
            {
                PageNumber = 1,
                FullText = "PART NO: 12345\nDESCRIPTION: MOUNTING BRACKET\nREV A",
                Width = 612,
                Height = 792
            };

            var result = _scanner.AnalyzePage(page, "/test.pdf");
            Assert.Equal("MOUNTING BRACKET", result.Description);
        }

        [Fact]
        public void AnalyzePage_WithRevision_ExtractsRevision()
        {
            var page = new PageText
            {
                PageNumber = 1,
                FullText = "PART NO: 12345\nREVISION: C",
                Width = 612,
                Height = 792
            };

            var result = _scanner.AnalyzePage(page, "/test.pdf");
            Assert.Equal("C", result.Revision);
        }

        [Fact]
        public void AnalyzePage_WithDeburNote_ExtractsNote()
        {
            var page = new PageText
            {
                PageNumber = 1,
                FullText = "PART NO: 12345\nNOTES:\n1. BREAK ALL EDGES\n2. DEBURR ALL",
                Width = 612,
                Height = 792
            };

            var result = _scanner.AnalyzePage(page, "/test.pdf");
            Assert.True(result.Notes.Count > 0);
            Assert.Contains(result.Notes, n => n.Category == NoteCategory.Deburr);
        }

        [Fact]
        public void AnalyzePage_WithRoutingHints_GeneratesHints()
        {
            var page = new PageText
            {
                PageNumber = 1,
                FullText = "PART NO: 12345\nBREAK ALL EDGES\nPOWDER COAT RED",
                Width = 612,
                Height = 792
            };

            var result = _scanner.AnalyzePage(page, "/test.pdf");
            Assert.True(result.RoutingHints.Count > 0);
        }

        [Fact]
        public void AnalyzePage_WithBomHeader_DetectsBom()
        {
            var page = new PageText
            {
                PageNumber = 1,
                FullText = "PART NO: ASSY-100\nDESCRIPTION: WELDMENT ASSEMBLY\n\n" +
                           "BILL OF MATERIAL\n" +
                           "ITEM NO  PART NUMBER  DESCRIPTION  QTY\n" +
                           "1 12345-01 BRACKET 2\n" +
                           "2 12345-02 PLATE 1\n",
                Width = 612,
                Height = 792
            };

            var result = _scanner.AnalyzePage(page, "/test.pdf");
            Assert.True(result.HasBom);
            Assert.True(result.IsAssemblyLevel);
        }

        [Fact]
        public void AnalyzePage_WithAssemblyKeyword_DetectsAssembly()
        {
            var page = new PageText
            {
                PageNumber = 1,
                FullText = "PART NO: ASSY-100\nDESCRIPTION: WELDED ASSY\nREV A",
                Width = 612,
                Height = 792
            };

            var result = _scanner.AnalyzePage(page, "/test.pdf");
            Assert.True(result.IsAssemblyLevel);
        }

        [Fact]
        public void AnalyzePage_RegularPart_NotAssemblyLevel()
        {
            var page = new PageText
            {
                PageNumber = 1,
                FullText = "PART NO: 12345-01\nDESCRIPTION: BRACKET\nREV A\nMATERIAL: A36",
                Width = 612,
                Height = 792
            };

            var result = _scanner.AnalyzePage(page, "/test.pdf");
            Assert.False(result.IsAssemblyLevel);
        }

        [Fact]
        public void AnalyzePage_ConfidenceIsPositive()
        {
            var page = new PageText
            {
                PageNumber = 1,
                FullText = "PART NO: 12345\nDESCRIPTION: BRACKET\nMATERIAL: A36",
                Width = 612,
                Height = 792
            };

            var result = _scanner.AnalyzePage(page, "/test.pdf");
            Assert.True(result.Confidence > 0);
        }

        // --- Multi-page grouping ---

        [Fact]
        public void ScanSinglePdf_MultiplePages_GroupsByPartNumber()
        {
            // This tests the grouping logic using AnalyzePage directly
            // (since we can't easily create real multi-page PDFs in tests)
            var index = new DrawingPackageIndex();

            var page1 = new PageText
            {
                PageNumber = 1,
                FullText = "PART NO: 12345-01\nDESCRIPTION: BRACKET\nSHEET 1 OF 2",
                Width = 612,
                Height = 792
            };

            var page2 = new PageText
            {
                PageNumber = 2,
                FullText = "PART NO: 12345-01\nSHEET 2 OF 2\nNOTES:\n1. BREAK ALL EDGES",
                Width = 612,
                Height = 792
            };

            var page3 = new PageText
            {
                PageNumber = 3,
                FullText = "PART NO: 67890\nDESCRIPTION: PLATE\nREV B",
                Width = 612,
                Height = 792
            };

            // Simulate what ScanSinglePdf does internally
            foreach (var page in new[] { page1, page2, page3 })
            {
                var pageInfo = _scanner.AnalyzePage(page, "/package.pdf");
                index.TotalPages++;

                if (!string.IsNullOrWhiteSpace(pageInfo.PartNumber))
                {
                    if (!index.PagesByPartNumber.TryGetValue(pageInfo.PartNumber, out var list))
                    {
                        list = new System.Collections.Generic.List<DrawingPageInfo>();
                        index.PagesByPartNumber[pageInfo.PartNumber] = list;
                    }
                    list.Add(pageInfo);
                }
                else
                {
                    index.UnmatchedPages.Add(pageInfo);
                }
            }

            Assert.Equal(3, index.TotalPages);
            Assert.Equal(2, index.UniquePartNumbers);
            Assert.Equal(3, index.MatchedPages);
            Assert.Empty(index.UnmatchedPages);

            // Part 12345-01 should have 2 pages
            Assert.True(index.PagesByPartNumber.ContainsKey("12345-01"));
            Assert.Equal(2, index.PagesByPartNumber["12345-01"].Count);

            // Part 67890 should have 1 page
            Assert.True(index.PagesByPartNumber.ContainsKey("67890"));
            Assert.Single(index.PagesByPartNumber["67890"]);
        }
    }
}
