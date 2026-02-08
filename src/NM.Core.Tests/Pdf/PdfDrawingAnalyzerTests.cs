using Xunit;
using NM.Core.Pdf;
using NM.Core.Pdf.Models;

namespace NM.Core.Tests.Pdf
{
    public class PdfDrawingAnalyzerTests
    {
        [Fact]
        public void FindCompanionPdf_ReturnsNull_ForNullPath()
        {
            Assert.Null(PdfDrawingAnalyzer.FindCompanionPdf(null));
        }

        [Fact]
        public void FindCompanionPdf_ReturnsNull_ForEmptyPath()
        {
            Assert.Null(PdfDrawingAnalyzer.FindCompanionPdf(""));
        }

        [Fact]
        public void FindCompanionPdf_ReturnsNull_WhenNoMatchingPdf()
        {
            // Path to a file that exists but has no companion PDF
            Assert.Null(PdfDrawingAnalyzer.FindCompanionPdf(@"C:\nonexistent\part.sldprt"));
        }

        [Fact]
        public void Analyze_ReturnsEmptyData_ForNonexistentFile()
        {
            var analyzer = new PdfDrawingAnalyzer();
            var result = analyzer.Analyze(@"C:\nonexistent\drawing.pdf");

            Assert.NotNull(result);
            Assert.Equal(0, result.OverallConfidence);
            Assert.Equal(0, result.PageCount);
        }

        [Fact]
        public void Analyze_ReturnsEmptyData_ForNullPath()
        {
            var analyzer = new PdfDrawingAnalyzer();
            var result = analyzer.Analyze(null);

            Assert.NotNull(result);
            Assert.Equal(0, result.OverallConfidence);
        }

        [Fact]
        public void DrawingData_DefaultValues()
        {
            var data = new DrawingData();

            Assert.Null(data.PartNumber);
            Assert.Null(data.Material);
            Assert.Null(data.Revision);
            Assert.NotNull(data.Notes);
            Assert.Empty(data.Notes);
            Assert.NotNull(data.RoutingHints);
            Assert.Empty(data.RoutingHints);
            Assert.NotNull(data.BomEntries);
            Assert.Empty(data.BomEntries);
            Assert.NotNull(data.GdtCallouts);
            Assert.Empty(data.GdtCallouts);
            Assert.Equal(AnalysisMethod.TextOnly, data.Method);
        }

        [Fact]
        public void TitleBlockInfo_OverallConfidence_CalculatesCorrectly()
        {
            var info = new TitleBlockInfo
            {
                PartNumber = "12345",
                PartNumberConfidence = 0.9,
                Material = "304 SS",
                MaterialConfidence = 0.8,
                Revision = "A",
                RevisionConfidence = 0.95
            };

            // Average of the 3 populated fields
            double expected = (0.9 + 0.8 + 0.95) / 3.0;
            Assert.Equal(expected, info.OverallConfidence, 4);
        }

        [Fact]
        public void TitleBlockInfo_OverallConfidence_ZeroWhenEmpty()
        {
            var info = new TitleBlockInfo();
            Assert.Equal(0, info.OverallConfidence);
        }

        [Fact]
        public void RoutingHint_DefaultValues()
        {
            var hint = new RoutingHint();
            Assert.Null(hint.WorkCenter);
            Assert.Null(hint.NoteText);
            Assert.Equal(0, hint.Confidence);
        }

        [Fact]
        public void DrawingNote_ConstructorSetsImpact()
        {
            var note = new DrawingNote("BREAK ALL EDGES", NoteCategory.Deburr, 0.95);
            Assert.Equal(RoutingImpact.AddOperation, note.Impact);
            Assert.Equal(0.95, note.Confidence);
        }

        [Fact]
        public void DrawingNote_GeneralCategory_IsInformational()
        {
            var note = new DrawingNote("Some general note", NoteCategory.General);
            Assert.Equal(RoutingImpact.Informational, note.Impact);
        }

        [Fact]
        public void DrawingNote_ProcessConstraint_IsModifyOperation()
        {
            var note = new DrawingNote("WATERJET ONLY", NoteCategory.ProcessConstraint);
            Assert.Equal(RoutingImpact.ModifyOperation, note.Impact);
        }
    }
}
