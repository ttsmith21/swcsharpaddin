using Xunit;
using NM.Core.AI.Models;
using NM.Core.Pdf;
using NM.Core.Pdf.Models;

namespace NM.Core.Tests.AI
{
    public class MergeVisionResultsTests
    {
        [Fact]
        public void MergeVisionResults_FillsGaps()
        {
            var analyzer = new PdfDrawingAnalyzer();
            var target = new DrawingData
            {
                PartNumber = "12345",
                Material = "304 SS"
                // Description and Revision are empty
            };

            var vision = new VisionAnalysisResult
            {
                Success = true,
                PartNumber = new FieldResult("12345", 0.9),
                Description = new FieldResult("MOUNTING BRACKET", 0.85),
                Revision = new FieldResult("C", 0.9),
                Material = new FieldResult("304 STAINLESS", 0.9)
            };

            analyzer.MergeVisionResults(target, vision);

            // Existing values preserved
            Assert.Equal("12345", target.PartNumber);
            Assert.Equal("304 SS", target.Material);  // Original kept, not overwritten

            // Gaps filled
            Assert.Equal("MOUNTING BRACKET", target.Description);
            Assert.Equal("C", target.Revision);
        }

        [Fact]
        public void MergeVisionResults_DoesNotOverwriteExistingValues()
        {
            var analyzer = new PdfDrawingAnalyzer();
            var target = new DrawingData
            {
                PartNumber = "ORIGINAL-PN",
                Description = "ORIGINAL DESC"
            };

            var vision = new VisionAnalysisResult
            {
                Success = true,
                PartNumber = new FieldResult("AI-PN", 0.95),
                Description = new FieldResult("AI DESC", 0.95)
            };

            analyzer.MergeVisionResults(target, vision);

            Assert.Equal("ORIGINAL-PN", target.PartNumber);
            Assert.Equal("ORIGINAL DESC", target.Description);
        }

        [Fact]
        public void MergeVisionResults_MergesNewNotes()
        {
            var analyzer = new PdfDrawingAnalyzer();
            var target = new DrawingData();
            target.Notes.Add(new DrawingNote("BREAK ALL EDGES", NoteCategory.Deburr, 0.95));

            var vision = new VisionAnalysisResult { Success = true };
            vision.ManufacturingNotes.Add(new NoteResult
            {
                Text = "BREAK ALL EDGES",  // Duplicate — should not be added
                Category = "deburr",
                Confidence = 0.9
            });
            vision.ManufacturingNotes.Add(new NoteResult
            {
                Text = "GALVANIZE PER ASTM A123",  // New — should be added
                Category = "finish",
                Confidence = 0.85
            });

            analyzer.MergeVisionResults(target, vision);

            Assert.Equal(2, target.Notes.Count);
            Assert.Contains(target.Notes, n => n.Text == "GALVANIZE PER ASTM A123");
        }

        [Fact]
        public void MergeVisionResults_SetsHybridMethod()
        {
            var analyzer = new PdfDrawingAnalyzer();
            var target = new DrawingData { RawText = "some text" };
            var vision = new VisionAnalysisResult { Success = true };

            analyzer.MergeVisionResults(target, vision);
            Assert.Equal(AnalysisMethod.Hybrid, target.Method);
        }

        [Fact]
        public void MergeVisionResults_SetsVisionMethodWhenNoText()
        {
            var analyzer = new PdfDrawingAnalyzer();
            var target = new DrawingData { RawText = null };
            var vision = new VisionAnalysisResult { Success = true };

            analyzer.MergeVisionResults(target, vision);
            Assert.Equal(AnalysisMethod.VisionAI, target.Method);
        }

        [Fact]
        public void MergeVisionResults_IgnoresFailedVision()
        {
            var analyzer = new PdfDrawingAnalyzer();
            var target = new DrawingData { PartNumber = null };
            var vision = new VisionAnalysisResult
            {
                Success = false,
                PartNumber = new FieldResult("SHOULD-NOT-APPLY", 0.9)
            };

            analyzer.MergeVisionResults(target, vision);
            Assert.Null(target.PartNumber);
        }

        [Fact]
        public void MergeVisionResults_BoostsConfidence()
        {
            var analyzer = new PdfDrawingAnalyzer();
            var target = new DrawingData { OverallConfidence = 0.5 };
            var vision = new VisionAnalysisResult
            {
                Success = true,
                PartNumber = new FieldResult("123", 0.95),
                Material = new FieldResult("A36", 0.9)
            };

            analyzer.MergeVisionResults(target, vision);
            Assert.True(target.OverallConfidence > 0.5);
        }
    }
}
