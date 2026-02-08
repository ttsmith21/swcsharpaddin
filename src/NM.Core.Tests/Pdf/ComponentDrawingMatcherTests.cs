using System.Collections.Generic;
using Xunit;
using NM.Core.Pdf;
using NM.Core.Pdf.Models;

namespace NM.Core.Tests.Pdf
{
    public class ComponentDrawingMatcherTests
    {
        private readonly ComponentDrawingMatcher _matcher = new ComponentDrawingMatcher();

        private static DrawingPackageIndex BuildIndex(params (string partNumber, string pdfPath)[] entries)
        {
            var index = new DrawingPackageIndex();
            foreach (var (pn, pdf) in entries)
            {
                var page = new DrawingPageInfo
                {
                    PartNumber = pn,
                    PdfPath = pdf,
                    PageNumber = 1,
                    HasText = true,
                    Confidence = 0.85
                };

                if (!index.PagesByPartNumber.TryGetValue(pn, out var list))
                {
                    list = new List<DrawingPageInfo>();
                    index.PagesByPartNumber[pn] = list;
                }
                list.Add(page);
            }
            return index;
        }

        // --- Single component matching ---

        [Fact]
        public void Match_ExactPartNumber_ReturnsExactMatch()
        {
            var index = BuildIndex(("12345-01", "/drawings/12345-01.pdf"));

            var match = _matcher.Match("/parts/12345-01.SLDPRT", "12345-01", index);

            Assert.True(match.IsMatched);
            Assert.Equal(MatchMethod.ExactPartNumber, match.Method);
            Assert.Equal(0.95, match.Confidence);
            Assert.Single(match.Pages);
        }

        [Fact]
        public void Match_CaseInsensitivePartNumber_Matches()
        {
            var index = BuildIndex(("12345-01", "/drawings/test.pdf"));

            var match = _matcher.Match("/parts/test.SLDPRT", "12345-01", index);

            Assert.True(match.IsMatched);
            Assert.Equal(MatchMethod.ExactPartNumber, match.Method);
        }

        [Fact]
        public void Match_ByFileName_WhenNoPartNumber()
        {
            var index = BuildIndex(("12345-01", "/drawings/12345-01.pdf"));

            var match = _matcher.Match("/parts/12345-01.SLDPRT", null, index);

            Assert.True(match.IsMatched);
            Assert.Equal(MatchMethod.FileName, match.Method);
            Assert.Equal(0.85, match.Confidence);
        }

        [Fact]
        public void Match_ByFileName_WhenPartNumberDoesntMatch()
        {
            var index = BuildIndex(("12345-01", "/drawings/12345-01.pdf"));

            var match = _matcher.Match("/parts/12345-01.SLDPRT", "WRONG-PN", index);

            Assert.True(match.IsMatched);
            Assert.Equal(MatchMethod.FileName, match.Method);
        }

        [Fact]
        public void Match_NoMatch_ReturnsNoMatch()
        {
            var index = BuildIndex(("12345-01", "/drawings/12345-01.pdf"));

            var match = _matcher.Match("/parts/99999.SLDPRT", "99999", index);

            Assert.False(match.IsMatched);
            Assert.Equal(MatchMethod.None, match.Method);
        }

        [Fact]
        public void Match_NullIndex_ReturnsNoMatch()
        {
            var match = _matcher.Match("/parts/test.SLDPRT", "12345", null);
            Assert.False(match.IsMatched);
        }

        [Fact]
        public void Match_NullPath_StillTriesPartNumber()
        {
            var index = BuildIndex(("12345-01", "/drawings/12345-01.pdf"));

            var match = _matcher.Match(null, "12345-01", index);

            Assert.True(match.IsMatched);
            Assert.Equal(MatchMethod.ExactPartNumber, match.Method);
        }

        [Fact]
        public void Match_PartNumberPreferred_OverFileName()
        {
            var index = BuildIndex(
                ("PN-FROM-PROPS", "/drawings/pn.pdf"),
                ("BRACKET", "/drawings/bracket.pdf"));

            var match = _matcher.Match("/parts/BRACKET.SLDPRT", "PN-FROM-PROPS", index);

            // Part number match should take precedence
            Assert.True(match.IsMatched);
            Assert.Equal(MatchMethod.ExactPartNumber, match.Method);
            Assert.Equal("/drawings/pn.pdf", match.Pages[0].PdfPath);
        }

        // --- BOM reference matching ---

        [Fact]
        public void Match_ViaBom_WhenPartInBomAndHasPage()
        {
            var index = BuildIndex(("12345-01", "/drawings/12345-01.pdf"));
            index.AllBomEntries.Add(new BomEntry
            {
                ItemNumber = "1",
                PartNumber = "12345-01",
                Description = "BRACKET",
                Quantity = 2,
                Confidence = 0.70
            });

            // Part number not in index directly, but file name doesn't match either
            // However, BOM reference exists
            var match = _matcher.Match("/parts/unknown.SLDPRT", "12345-01", index);

            // Should find via exact part number first
            Assert.True(match.IsMatched);
            Assert.Equal(MatchMethod.ExactPartNumber, match.Method);
        }

        [Fact]
        public void Match_ViaBom_WhenOnlyBomReferenceExists()
        {
            var index = new DrawingPackageIndex();

            // Add a page with a different part number
            var page = new DrawingPageInfo
            {
                PartNumber = "BRACKET-A",
                PdfPath = "/drawings/bracket.pdf",
                PageNumber = 1,
                HasText = true
            };
            index.PagesByPartNumber["BRACKET-A"] = new List<DrawingPageInfo> { page };

            // BOM entry links "BRACKET-A" to the drawing
            index.AllBomEntries.Add(new BomEntry
            {
                PartNumber = "BRACKET-A",
                Quantity = 2
            });

            var match = _matcher.Match("/parts/BRACKET-A.SLDPRT", null, index);

            Assert.True(match.IsMatched);
        }

        // --- MatchAll ---

        [Fact]
        public void MatchAll_MultipleComponents_ReturnsMatchedAndUnmatched()
        {
            var index = BuildIndex(
                ("12345-01", "/drawings/12345-01.pdf"),
                ("12345-02", "/drawings/12345-02.pdf"),
                ("12345-03", "/drawings/12345-03.pdf"));

            var components = new List<ComponentInfo>
            {
                new ComponentInfo { FilePath = "/parts/12345-01.SLDPRT", PartNumber = "12345-01" },
                new ComponentInfo { FilePath = "/parts/12345-02.SLDPRT", PartNumber = "12345-02" },
                new ComponentInfo { FilePath = "/parts/99999.SLDPRT", PartNumber = "99999" },
            };

            var results = _matcher.MatchAll(components, index);

            Assert.Equal(2, results.Matched.Count);
            Assert.Single(results.Unmatched);
            Assert.Equal("/parts/99999.SLDPRT", results.Unmatched[0]);
        }

        [Fact]
        public void MatchAll_IdentifiesUnmatchedDrawings()
        {
            var index = BuildIndex(
                ("12345-01", "/drawings/12345-01.pdf"),
                ("12345-02", "/drawings/12345-02.pdf"),
                ("ORPHAN-DWG", "/drawings/orphan.pdf"));

            var components = new List<ComponentInfo>
            {
                new ComponentInfo { FilePath = "/parts/12345-01.SLDPRT", PartNumber = "12345-01" },
                new ComponentInfo { FilePath = "/parts/12345-02.SLDPRT", PartNumber = "12345-02" },
            };

            var results = _matcher.MatchAll(components, index);

            Assert.Equal(2, results.Matched.Count);
            Assert.Empty(results.Unmatched);
            Assert.Single(results.UnmatchedDrawings);
            Assert.Equal("ORPHAN-DWG", results.UnmatchedDrawings[0].PartNumber);
        }

        [Fact]
        public void MatchAll_NullComponents_ReturnsEmptyResults()
        {
            var index = BuildIndex(("12345", "/test.pdf"));
            var results = _matcher.MatchAll(null, index);
            Assert.Empty(results.Matched);
        }

        [Fact]
        public void MatchAll_NullIndex_ReturnsEmptyResults()
        {
            var components = new List<ComponentInfo>
            {
                new ComponentInfo { FilePath = "/parts/test.SLDPRT", PartNumber = "12345" }
            };
            var results = _matcher.MatchAll(components, null);
            Assert.Empty(results.Matched);
        }

        [Fact]
        public void MatchAll_IncludesUnmatchedPagesInUnmatchedDrawings()
        {
            var index = BuildIndex(("12345-01", "/drawings/12345-01.pdf"));
            index.UnmatchedPages.Add(new DrawingPageInfo
            {
                PdfPath = "/drawings/mystery.pdf",
                PageNumber = 3,
                HasText = true
            });

            var components = new List<ComponentInfo>
            {
                new ComponentInfo { FilePath = "/parts/12345-01.SLDPRT", PartNumber = "12345-01" },
            };

            var results = _matcher.MatchAll(components, index);

            Assert.Single(results.Matched);
            // The unmatched page should appear in UnmatchedDrawings
            Assert.Single(results.UnmatchedDrawings);
        }

        // --- DrawingPackageIndex integration ---

        [Fact]
        public void Match_FuzzyPartNumber_VisFindPages()
        {
            // DrawingPackageIndex.FindPages does fuzzy matching
            var index = BuildIndex(("12345-01-REV-C", "/drawings/test.pdf"));

            // "12345-01-REV-C" contains "12345-01" — partial match
            var match = _matcher.Match("/parts/test.SLDPRT", "12345-01", index);

            Assert.True(match.IsMatched);
        }

        [Fact]
        public void Match_MultiplePages_ReturnsList()
        {
            var index = new DrawingPackageIndex();
            var pages = new List<DrawingPageInfo>
            {
                new DrawingPageInfo { PartNumber = "12345", PdfPath = "/a.pdf", PageNumber = 1, HasText = true },
                new DrawingPageInfo { PartNumber = "12345", PdfPath = "/a.pdf", PageNumber = 2, HasText = true }
            };
            index.PagesByPartNumber["12345"] = pages;

            var match = _matcher.Match("/parts/test.SLDPRT", "12345", index);

            Assert.True(match.IsMatched);
            Assert.Equal(2, match.Pages.Count);
        }
    }
}
