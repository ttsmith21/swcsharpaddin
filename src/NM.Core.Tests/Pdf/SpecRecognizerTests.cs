using System.Linq;
using Xunit;
using NM.Core.Pdf;
using NM.Core.Pdf.Models;

namespace NM.Core.Tests.Pdf
{
    public class SpecRecognizerTests
    {
        private readonly SpecRecognizer _recognizer = new SpecRecognizer();

        // --- Database ---

        [Fact]
        public void DatabaseSize_IsSubstantial()
        {
            // Ensure we have a meaningful spec database
            Assert.True(SpecRecognizer.DatabaseSize >= 60,
                $"Expected at least 60 specs, got {SpecRecognizer.DatabaseSize}");
        }

        // --- Material specs ---

        [Theory]
        [InlineData("MATERIAL: ASTM A36", "ASTM A36", SpecCategory.Material)]
        [InlineData("ASTM A-36 STEEL", "ASTM A36", SpecCategory.Material)]
        [InlineData("PER ASTM A500 GRADE B", "ASTM A500", SpecCategory.Material)]
        [InlineData("ASTM A572 GR 50", "ASTM A572", SpecCategory.Material)]
        [InlineData("ASTM A240 TYPE 304", "ASTM A240", SpecCategory.Material)]
        public void Recognize_MaterialSpecs(string text, string expectedSpecId, SpecCategory expectedCategory)
        {
            var matches = _recognizer.Recognize(text);
            Assert.NotEmpty(matches);
            var match = matches.First();
            Assert.Equal(expectedSpecId, match.SpecId);
            Assert.Equal(expectedCategory, match.Category);
        }

        [Theory]
        [InlineData("AMS 4027", "AMS 4027", "6061-T6 Aluminum Sheet")]
        [InlineData("AMS 5510", "AMS 5510", "304 Stainless Sheet")]
        [InlineData("AMS 4911", "AMS 4911", "Titanium 6Al-4V Sheet/Plate")]
        [InlineData("AMS 6350", "AMS 6350", "4130 Normalized Steel")]
        public void Recognize_AerospaceAmsSpecs(string text, string expectedSpecId, string expectedName)
        {
            var matches = _recognizer.Recognize(text);
            Assert.NotEmpty(matches);
            var match = matches.First();
            Assert.Equal(expectedSpecId, match.SpecId);
            Assert.Contains(expectedName, match.FullName);
        }

        // --- Welding codes ---

        [Theory]
        [InlineData("WELD PER AWS D1.1", "AWS D1.1")]
        [InlineData("AWS D1.2 ALUMINUM", "AWS D1.2")]
        [InlineData("PER AWS D1.6", "AWS D1.6")]
        [InlineData("AWS D17.1 AEROSPACE", "AWS D17.1")]
        public void Recognize_WeldingCodes(string text, string expectedSpecId)
        {
            var matches = _recognizer.Recognize(text);
            Assert.NotEmpty(matches);
            var match = matches.First();
            Assert.Equal(expectedSpecId, match.SpecId);
            Assert.Equal(SpecCategory.Welding, match.Category);
            Assert.Equal(RoutingOp.Weld, match.RoutingOp);
            Assert.Equal("F400", match.WorkCenter);
        }

        // --- Coating / Paint specs ---

        [Theory]
        [InlineData("PRIME PER MIL-PRF-22750", "MIL-PRF-22750")]
        [InlineData("PAINT PER MIL-PRF-85285", "MIL-PRF-85285")]
        [InlineData("MIL-DTL-53039 CARC", "MIL-DTL-53039")]
        public void Recognize_CoatingSpecs(string text, string expectedSpecId)
        {
            var matches = _recognizer.Recognize(text);
            Assert.NotEmpty(matches);
            var match = matches.First();
            Assert.Equal(expectedSpecId, match.SpecId);
            Assert.Equal(SpecCategory.Coating, match.Category);
            Assert.Equal(RoutingOp.OutsideProcess, match.RoutingOp);
        }

        // --- Plating specs ---

        [Theory]
        [InlineData("ZINC PLATE PER ASTM B633", "ASTM B633")]
        [InlineData("PER MIL-DTL-5541 CLASS 1A", "MIL-DTL-5541")]
        [InlineData("BLACK OXIDE PER MIL-DTL-13924", "MIL-DTL-13924")]
        [InlineData("PASSIVATE PER AMS 2700", "AMS 2700")]
        [InlineData("PASSIVATE PER ASTM A967", "ASTM A967")]
        public void Recognize_PlatingSpecs(string text, string expectedSpecId)
        {
            var matches = _recognizer.Recognize(text);
            Assert.NotEmpty(matches);
            var match = matches.First();
            Assert.Equal(expectedSpecId, match.SpecId);
            Assert.Equal(RoutingOp.OutsideProcess, match.RoutingOp);
        }

        // --- Heat treat specs ---

        [Theory]
        [InlineData("HEAT TREAT PER AMS 2759", "AMS 2759")]
        [InlineData("PER AMS-H-6875", "AMS-H-6875")]
        public void Recognize_HeatTreatSpecs(string text, string expectedSpecId)
        {
            var matches = _recognizer.Recognize(text);
            Assert.NotEmpty(matches);
            var match = matches.First();
            Assert.Equal(expectedSpecId, match.SpecId);
            Assert.Equal(SpecCategory.HeatTreat, match.Category);
        }

        // --- Surface finish specs ---

        [Theory]
        [InlineData("SHOT PEEN PER AMS 2430", "AMS 2430")]
        [InlineData("ANODIZE PER MIL-A-8625 TYPE III", "MIL-A-8625")]
        [InlineData("AMS 2471 SULFURIC ANODIZE", "AMS 2471")]
        [InlineData("HARD ANODIZE PER AMS 2472", "AMS 2472")]
        public void Recognize_SurfaceFinishSpecs(string text, string expectedSpecId)
        {
            var matches = _recognizer.Recognize(text);
            Assert.NotEmpty(matches);
            var match = matches.First();
            Assert.Equal(expectedSpecId, match.SpecId);
            Assert.Equal(RoutingOp.OutsideProcess, match.RoutingOp);
        }

        // --- Inspection / Testing specs ---

        [Theory]
        [InlineData("DIMENSIONING PER ASME Y14.5", "ASME Y14.5")]
        [InlineData("FAI PER AS 9102", "AS 9102")]
        [InlineData("MAG PARTICLE PER ASTM E1444", "ASTM E1444")]
        [InlineData("LPI PER ASTM E1417", "ASTM E1417")]
        public void Recognize_InspectionSpecs(string text, string expectedSpecId)
        {
            var matches = _recognizer.Recognize(text);
            Assert.NotEmpty(matches);
            Assert.Equal(expectedSpecId, matches.First().SpecId);
        }

        [Fact]
        public void Recognize_AS9102_HasInspectRouting()
        {
            var matches = _recognizer.Recognize("FIRST ARTICLE PER AS 9102");
            Assert.NotEmpty(matches);
            Assert.Equal(RoutingOp.Inspect, matches.First().RoutingOp);
        }

        // --- Quality / Controlled ---

        [Theory]
        [InlineData("NADCAP CERTIFIED PROCESSOR REQUIRED", "NADCAP")]
        [InlineData("AS 9100 QUALITY SYSTEM", "AS 9100")]
        [InlineData("ISO 9001 CERTIFIED", "ISO 9001")]
        public void Recognize_QualityStandards(string text, string expectedSpecId)
        {
            var matches = _recognizer.Recognize(text);
            Assert.NotEmpty(matches);
            Assert.Equal(expectedSpecId, matches.First().SpecId);
            Assert.Equal(SpecCategory.Quality, matches.First().Category);
        }

        [Theory]
        [InlineData("THIS DOCUMENT CONTAINS ITAR CONTROLLED DATA", "ITAR")]
        [InlineData("DFARS 252.225-7009", "DFARS")]
        [InlineData("CUI MARKING REQUIRED", "CUI")]
        [InlineData("NIST 800-171 COMPLIANT", "NIST 800-171")]
        public void Recognize_ControlledFlags(string text, string expectedSpecId)
        {
            var matches = _recognizer.Recognize(text);
            Assert.NotEmpty(matches);
            Assert.Equal(expectedSpecId, matches.First().SpecId);
            Assert.Equal(SpecCategory.Controlled, matches.First().Category);
        }

        // --- Multiple specs in one text ---

        [Fact]
        public void Recognize_MultipleSpecs_ReturnsAll()
        {
            string text = @"
                MATERIAL: ASTM A36 STEEL
                WELD PER AWS D1.1
                PAINT PER MIL-PRF-85285
                BREAK ALL EDGES
                FIRST ARTICLE PER AS 9102
            ";

            var matches = _recognizer.Recognize(text);
            Assert.True(matches.Count >= 4, $"Expected >= 4, got {matches.Count}");

            var specIds = matches.Select(m => m.SpecId).ToList();
            Assert.Contains("ASTM A36", specIds);
            Assert.Contains("AWS D1.1", specIds);
            Assert.Contains("MIL-PRF-85285", specIds);
            Assert.Contains("AS 9102", specIds);
        }

        [Fact]
        public void Recognize_Deduplicates()
        {
            string text = "ASTM A36 ASTM A36 ASTM A36";
            var matches = _recognizer.Recognize(text);
            Assert.Single(matches);
        }

        // --- Empty / null ---

        [Fact]
        public void Recognize_NullText_ReturnsEmpty()
        {
            Assert.Empty(_recognizer.Recognize(null));
        }

        [Fact]
        public void Recognize_EmptyText_ReturnsEmpty()
        {
            Assert.Empty(_recognizer.Recognize(""));
        }

        [Fact]
        public void Recognize_NoSpecs_ReturnsEmpty()
        {
            Assert.Empty(_recognizer.Recognize("BREAK ALL EDGES. DEBURR ALL. POWDER COAT RED."));
        }

        // --- ToRoutingHints ---

        [Fact]
        public void ToRoutingHints_WeldSpec_CreatesWeldHint()
        {
            var matches = _recognizer.Recognize("WELD PER AWS D1.1");
            var hints = _recognizer.ToRoutingHints(matches);

            Assert.NotEmpty(hints);
            var hint = hints.First();
            Assert.Equal(RoutingOp.Weld, hint.Operation);
            Assert.Equal("F400", hint.WorkCenter);
            Assert.Contains("AWS D1.1", hint.NoteText);
        }

        [Fact]
        public void ToRoutingHints_OutsideProcess_NoWorkCenter()
        {
            var matches = _recognizer.Recognize("ANODIZE PER MIL-A-8625");
            var hints = _recognizer.ToRoutingHints(matches);

            Assert.NotEmpty(hints);
            var hint = hints.First();
            Assert.Equal(RoutingOp.OutsideProcess, hint.Operation);
            Assert.Null(hint.WorkCenter);
            Assert.Contains("MIL-A-8625", hint.NoteText);
        }

        [Fact]
        public void ToRoutingHints_MaterialOnly_NoHint()
        {
            var matches = _recognizer.Recognize("ASTM A36 STEEL");
            var hints = _recognizer.ToRoutingHints(matches);

            // Material specs have no routing operation
            Assert.Empty(hints);
        }

        [Fact]
        public void ToRoutingHints_Deduplicates_SameOperation()
        {
            // Two coating specs should produce only one outside process hint
            var matches = _recognizer.Recognize("PRIME PER MIL-PRF-22750 THEN PAINT PER MIL-PRF-85285");
            var hints = _recognizer.ToRoutingHints(matches);

            // Both are OutsideProcess with null workCenter, so only one hint
            Assert.Single(hints);
        }

        [Fact]
        public void ToRoutingHints_NullMatches_ReturnsEmpty()
        {
            Assert.Empty(_recognizer.ToRoutingHints(null));
        }

        // --- ToDrawingNotes ---

        [Fact]
        public void ToDrawingNotes_ConvertsToNoteObjects()
        {
            var matches = _recognizer.Recognize("WELD PER AWS D1.1");
            var notes = _recognizer.ToDrawingNotes(matches, pageNumber: 3);

            Assert.NotEmpty(notes);
            var note = notes.First();
            Assert.Contains("AWS D1.1", note.Text);
            Assert.Equal(NoteCategory.Weld, note.Category);
            Assert.Equal(RoutingImpact.AddOperation, note.Impact);
            Assert.Equal(3, note.PageNumber);
        }

        [Fact]
        public void ToDrawingNotes_MaterialSpec_IsInformational()
        {
            var matches = _recognizer.Recognize("ASTM A36");
            var notes = _recognizer.ToDrawingNotes(matches);

            Assert.NotEmpty(notes);
            Assert.Equal(RoutingImpact.Informational, notes.First().Impact);
        }

        [Fact]
        public void ToDrawingNotes_ControlledFlag_IsInformational()
        {
            var matches = _recognizer.Recognize("ITAR CONTROLLED");
            var notes = _recognizer.ToDrawingNotes(matches);

            Assert.NotEmpty(notes);
            Assert.Equal(RoutingImpact.Informational, notes.First().Impact);
        }

        // --- Confidence ---

        [Fact]
        public void Recognize_ControlledFlags_HighConfidence()
        {
            var matches = _recognizer.Recognize("ITAR");
            Assert.NotEmpty(matches);
            Assert.True(matches.First().Confidence >= 0.90);
        }

        [Fact]
        public void Recognize_WeldingCode_HighConfidence()
        {
            var matches = _recognizer.Recognize("AWS D1.1");
            Assert.NotEmpty(matches);
            Assert.True(matches.First().Confidence >= 0.90);
        }

        // --- Real-world drawing text ---

        [Fact]
        public void Recognize_RealisticDrawingText()
        {
            string drawingText = @"
                NORTHERN MANUFACTURING INC
                PART NO: 12345-01
                DESCRIPTION: SUPPORT BRACKET
                MATERIAL: ASTM A36 HOT ROLLED STEEL
                REV: C

                NOTES:
                1. BREAK ALL SHARP EDGES
                2. WELD PER AWS D1.1 ALL AROUND
                3. HOT DIP GALVANIZE PER ASTM A123
                4. UNLESS OTHERWISE SPECIFIED TOLERANCES PER ASME Y14.5

                THIS DRAWING IS ITAR CONTROLLED
                DFARS 252.225-7009 APPLIES
            ";

            var matches = _recognizer.Recognize(drawingText);

            var specIds = matches.Select(m => m.SpecId).ToList();
            Assert.Contains("ASTM A36", specIds);
            Assert.Contains("AWS D1.1", specIds);
            Assert.Contains("ASME Y14.5", specIds);
            Assert.Contains("ITAR", specIds);
            Assert.Contains("DFARS", specIds);
        }

        [Fact]
        public void Recognize_AerospaceDrawingText()
        {
            string text = @"
                AMS 4027 ALUMINUM 6061-T6
                ANODIZE PER MIL-A-8625 TYPE III CLASS 2
                SHOT PEEN PER AMS 2430
                FIRST ARTICLE INSPECTION PER AS 9102
                NADCAP CERTIFIED PROCESSOR REQUIRED
                FPI PER AMS 2644
                NIST 800-171 APPLIES
            ";

            var matches = _recognizer.Recognize(text);

            Assert.True(matches.Count >= 6, $"Expected >= 6 specs, got {matches.Count}");

            var specIds = matches.Select(m => m.SpecId).ToList();
            Assert.Contains("AMS 4027", specIds);
            Assert.Contains("MIL-A-8625", specIds);
            Assert.Contains("AMS 2430", specIds);
            Assert.Contains("AS 9102", specIds);
            Assert.Contains("NADCAP", specIds);
            Assert.Contains("AMS 2644", specIds);
        }
    }
}
