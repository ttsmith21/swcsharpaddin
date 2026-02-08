using System.Linq;
using Xunit;
using NM.Core.Pdf;
using NM.Core.Pdf.Models;

namespace NM.Core.Tests.Pdf
{
    public class DrawingNoteExtractorTests
    {
        private readonly DrawingNoteExtractor _extractor = new DrawingNoteExtractor();

        // --- Deburr notes ---

        [Theory]
        [InlineData("BREAK ALL EDGES")]
        [InlineData("break all sharp edges")]
        [InlineData("DEBURR ALL")]
        [InlineData("REMOVE ALL BURRS")]
        [InlineData("TUMBLE DEBURR")]
        public void DetectsDeburNotes(string note)
        {
            var notes = _extractor.ExtractNotes(note);
            Assert.Single(notes);
            Assert.Equal(NoteCategory.Deburr, notes[0].Category);
        }

        // --- Finish notes ---

        [Theory]
        [InlineData("POWDER COAT BLACK")]
        [InlineData("ANODIZE PER SPEC")]
        [InlineData("GALVANIZE")]
        [InlineData("ZINC PLATE")]
        [InlineData("CHROME PLATE")]
        [InlineData("BLACK OXIDE")]
        [InlineData("HOT DIP GALVANIZE")]
        public void DetectsFinishNotes(string note)
        {
            var notes = _extractor.ExtractNotes(note);
            Assert.NotEmpty(notes);
            Assert.Contains(notes, n => n.Category == NoteCategory.Finish);
        }

        // --- Heat treat notes ---

        [Theory]
        [InlineData("HEAT TREAT TO RC 60")]
        [InlineData("STRESS RELIEVE")]
        [InlineData("HARDEN TO HRC 55")]
        [InlineData("NORMALIZE")]
        [InlineData("CASE HARDEN")]
        public void DetectsHeatTreatNotes(string note)
        {
            var notes = _extractor.ExtractNotes(note);
            Assert.NotEmpty(notes);
            Assert.Contains(notes, n => n.Category == NoteCategory.HeatTreat);
        }

        // --- Weld notes ---

        [Theory]
        [InlineData("WELD ALL AROUND")]
        [InlineData("MIG WELD")]
        [InlineData("TIG WELD")]
        [InlineData("SPOT WELD")]
        [InlineData("FILLET WELD")]
        public void DetectsWeldNotes(string note)
        {
            var notes = _extractor.ExtractNotes(note);
            Assert.NotEmpty(notes);
            Assert.Contains(notes, n => n.Category == NoteCategory.Weld);
        }

        // --- Process constraint notes ---

        [Theory]
        [InlineData("WATERJET ONLY")]
        [InlineData("LASER CUT")]
        [InlineData("DO NOT LASER")]
        public void DetectsProcessConstraints(string note)
        {
            var notes = _extractor.ExtractNotes(note);
            Assert.NotEmpty(notes);
            Assert.Contains(notes, n => n.Category == NoteCategory.ProcessConstraint);
        }

        // --- Inspection notes ---

        [Theory]
        [InlineData("INSPECT PER DRAWING")]
        [InlineData("CMM INSPECT")]
        [InlineData("FIRST ARTICLE REQUIRED")]
        [InlineData("PPAP REQUIRED")]
        public void DetectsInspectionNotes(string note)
        {
            var notes = _extractor.ExtractNotes(note);
            Assert.NotEmpty(notes);
            Assert.Contains(notes, n => n.Category == NoteCategory.Inspect);
        }

        // --- Hardware notes ---

        [Theory]
        [InlineData("INSTALL PEM HARDWARE")]
        [InlineData("INSERT RIVET NUT")]
        public void DetectsHardwareNotes(string note)
        {
            var notes = _extractor.ExtractNotes(note);
            Assert.NotEmpty(notes);
            Assert.Contains(notes, n => n.Category == NoteCategory.Hardware);
        }

        // --- Routing hint generation ---

        [Fact]
        public void GeneratesDeburRoutingHint()
        {
            var notes = _extractor.ExtractNotes("BREAK ALL EDGES");
            var hints = _extractor.GenerateRoutingHints(notes);

            Assert.Single(hints);
            Assert.Equal(RoutingOp.Deburr, hints[0].Operation);
            Assert.Equal("F210", hints[0].WorkCenter);
            Assert.Equal("BREAK ALL EDGES", hints[0].NoteText);
        }

        [Fact]
        public void GeneratesWeldRoutingHint()
        {
            var notes = _extractor.ExtractNotes("WELD ALL AROUND");
            var hints = _extractor.GenerateRoutingHints(notes);

            Assert.Single(hints);
            Assert.Equal(RoutingOp.Weld, hints[0].Operation);
            Assert.Equal("F400", hints[0].WorkCenter);
        }

        [Fact]
        public void GeneratesOutsideProcessHint()
        {
            var notes = _extractor.ExtractNotes("GALVANIZE");
            var hints = _extractor.GenerateRoutingHints(notes);

            Assert.Single(hints);
            Assert.Equal(RoutingOp.OutsideProcess, hints[0].Operation);
            Assert.Null(hints[0].WorkCenter); // Outside process, no internal WC
        }

        [Fact]
        public void GeneratesProcessOverrideHint()
        {
            var notes = _extractor.ExtractNotes("WATERJET ONLY");
            var hints = _extractor.GenerateRoutingHints(notes);

            Assert.Single(hints);
            Assert.Equal(RoutingOp.ProcessOverride, hints[0].Operation);
            Assert.Equal("F110", hints[0].WorkCenter);
        }

        [Fact]
        public void MultipleNotes_GenerateMultipleHints()
        {
            string text = "BREAK ALL EDGES\nWELD ALL AROUND\nPOWDER COAT BLACK\nFIRST ARTICLE REQUIRED";
            var notes = _extractor.ExtractNotes(text);
            var hints = _extractor.GenerateRoutingHints(notes);

            Assert.True(notes.Count >= 4, $"Expected at least 4 notes, got {notes.Count}");
            Assert.True(hints.Count >= 4, $"Expected at least 4 routing hints, got {hints.Count}");

            Assert.Contains(hints, h => h.Operation == RoutingOp.Deburr);
            Assert.Contains(hints, h => h.Operation == RoutingOp.Weld);
            Assert.Contains(hints, h => h.Operation == RoutingOp.OutsideProcess);
            Assert.Contains(hints, h => h.Operation == RoutingOp.Inspect);
        }

        [Fact]
        public void EmptyText_ReturnsEmpty()
        {
            var notes = _extractor.ExtractNotes("");
            Assert.Empty(notes);
        }

        [Fact]
        public void DuplicateNotes_AreDeduped()
        {
            string text = "BREAK ALL EDGES\nBREAK ALL EDGES\nBREAK ALL EDGES";
            var notes = _extractor.ExtractNotes(text);
            Assert.Single(notes);
        }

        [Fact]
        public void AllNotesHaveConfidence()
        {
            string text = "BREAK ALL EDGES\nWELD PER DRAWING\nANODIZE";
            var notes = _extractor.ExtractNotes(text);
            Assert.All(notes, n => Assert.True(n.Confidence > 0));
        }
    }
}
