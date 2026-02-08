using System.Collections.Generic;
using System.Linq;
using Xunit;
using NM.Core.Pdf.Models;
using NM.Core.Reconciliation;
using NM.Core.Reconciliation.Models;

namespace NM.Core.Tests.Reconciliation
{
    public class PropertySuggestionServiceTests
    {
        private readonly PropertySuggestionService _service = new PropertySuggestionService();

        // --- Part identity suggestions ---

        [Fact]
        public void PartSuggestions_FillsDescription()
        {
            var recon = new ReconciliationResult();
            recon.GapFills.Add(new GapFill
            {
                Field = "Description",
                Value = "MOUNTING BRACKET",
                Source = "PDF title block",
                Confidence = 0.85
            });

            var suggestions = _service.GeneratePartSuggestions(recon, new Dictionary<string, string>());

            var desc = suggestions.First(s => s.PropertyName == "Description");
            Assert.Equal("MOUNTING BRACKET", desc.Value);
            Assert.Equal(PropertyCategory.Identity, desc.Category);
            Assert.True(desc.IsGapFill);
        }

        [Fact]
        public void PartSuggestions_FillsRevision()
        {
            var recon = new ReconciliationResult();
            recon.GapFills.Add(new GapFill { Field = "Revision", Value = "C", Source = "PDF title block", Confidence = 0.90 });

            var suggestions = _service.GeneratePartSuggestions(recon, new Dictionary<string, string>());

            Assert.Contains(suggestions, s => s.PropertyName == "Revision" && s.Value == "C");
        }

        [Fact]
        public void PartSuggestions_MapsMaterialToRbMaterialType()
        {
            var recon = new ReconciliationResult();
            recon.GapFills.Add(new GapFill { Field = "Material", Value = "304 SS", Source = "PDF title block", Confidence = 0.85 });

            var suggestions = _service.GeneratePartSuggestions(recon, new Dictionary<string, string>());

            // Material maps to "rbMaterialType" property name
            Assert.Contains(suggestions, s => s.PropertyName == "rbMaterialType" && s.Value == "304 SS");
        }

        // --- Part routing suggestions ---

        [Fact]
        public void PartSuggestions_DeburNote_MapsToF210()
        {
            var recon = new ReconciliationResult();
            recon.RoutingSuggestions.Add(new RoutingSuggestion
            {
                Operation = RoutingOp.Deburr,
                WorkCenter = "F210",
                NoteText = "BREAK ALL EDGES",
                SourceNote = "BREAK ALL EDGES",
                Confidence = 0.95
            });

            var suggestions = _service.GeneratePartSuggestions(recon, new Dictionary<string, string>());

            var note = suggestions.First(s => s.PropertyName == "F210_Note");
            Assert.Equal("BREAK ALL EDGES", note.Value);
            Assert.Equal(PropertyCategory.Routing, note.Category);
        }

        [Fact]
        public void PartSuggestions_TapNote_MapsToF220()
        {
            var recon = new ReconciliationResult();
            recon.RoutingSuggestions.Add(new RoutingSuggestion
            {
                Operation = RoutingOp.Tap,
                WorkCenter = "F220",
                NoteText = "TAP 1/4-20",
                SourceNote = "TAP 1/4-20",
                Confidence = 0.90
            });

            var suggestions = _service.GeneratePartSuggestions(recon, new Dictionary<string, string>());

            Assert.Contains(suggestions, s => s.PropertyName == "F220_Note" && s.Value.Contains("TAP 1/4-20"));
        }

        [Fact]
        public void PartSuggestions_ProcessOverride_MapsToOP20WorkCenter()
        {
            var recon = new ReconciliationResult();
            recon.RoutingSuggestions.Add(new RoutingSuggestion
            {
                Operation = RoutingOp.ProcessOverride,
                WorkCenter = "F110",
                NoteText = "WATERJET ONLY",
                SourceNote = "WATERJET ONLY",
                Confidence = 0.95
            });

            var suggestions = _service.GeneratePartSuggestions(recon, new Dictionary<string, string>());

            Assert.Contains(suggestions, s => s.PropertyName == "OP20_WorkCenter" && s.Value == "F110");
        }

        [Fact]
        public void PartSuggestions_OutsideProcess_AddsNote()
        {
            var recon = new ReconciliationResult();
            recon.RoutingSuggestions.Add(new RoutingSuggestion
            {
                Operation = RoutingOp.OutsideProcess,
                NoteText = "GALVANIZE",
                SourceNote = "GALVANIZE",
                Confidence = 0.90
            });

            var suggestions = _service.GeneratePartSuggestions(recon, new Dictionary<string, string>());

            Assert.Contains(suggestions, s => s.PropertyName == "OutsideProcessNote");
        }

        [Fact]
        public void PartSuggestions_AppendsToExistingNote()
        {
            var recon = new ReconciliationResult();
            recon.RoutingSuggestions.Add(new RoutingSuggestion
            {
                Operation = RoutingOp.Deburr,
                NoteText = "TUMBLE DEBURR",
                SourceNote = "TUMBLE DEBURR",
                Confidence = 0.90
            });

            var currentProps = new Dictionary<string, string>
            {
                { "F210_Note", "BREAK ALL EDGES" }
            };

            var suggestions = _service.GeneratePartSuggestions(recon, currentProps);

            var note = suggestions.First(s => s.PropertyName == "F210_Note");
            Assert.Contains("BREAK ALL EDGES", note.Value);
            Assert.Contains("TUMBLE DEBURR", note.Value);
            Assert.True(note.IsOverride); // Changing existing value
        }

        // --- Assembly operation suggestions ---

        [Fact]
        public void AssemblyOperations_GeneratesSlots()
        {
            var recon = new ReconciliationResult();
            recon.RoutingSuggestions.Add(new RoutingSuggestion
            {
                Operation = RoutingOp.Weld,
                WorkCenter = "F400",
                NoteText = "WELD PER DWG",
                SourceNote = "WELD ALL AROUND",
                Confidence = 0.90
            });
            recon.RoutingSuggestions.Add(new RoutingSuggestion
            {
                Operation = RoutingOp.OutsideProcess,
                NoteText = "POWDER COAT BLACK",
                SourceNote = "POWDER COAT BLACK",
                Confidence = 0.85
            });

            var operations = _service.GenerateAssemblyOperations(recon, startingOpNumber: 20);

            Assert.Equal(2, operations.Count);

            // First operation: Weld at OP20
            Assert.Equal(20, operations[0].OpNumber);
            Assert.Equal("F400", operations[0].WorkCenter);
            Assert.Equal("WELD PER DWG", operations[0].RoutingNote);

            // Second operation: Outside process at OP30
            Assert.Equal(30, operations[1].OpNumber);
            Assert.Equal("POWDER COAT BLACK", operations[1].RoutingNote);
        }

        [Fact]
        public void AssemblyOperations_ToProperties_GeneratesCorrectKeys()
        {
            var op = new AssemblyOperationSuggestion
            {
                OpNumber = 30,
                WorkCenter = "F400",
                Setup_min = 15,
                Run_min = 10,
                RoutingNote = "WELD PER DWG"
            };

            var props = op.ToProperties();

            Assert.Equal("F400", props["OP30_WC"]);
            Assert.Equal("15", props["OP30_S"]);
            Assert.Equal("10", props["OP30_R"]);
            Assert.Equal("WELD PER DWG", props["OP30_RN"]);
        }

        [Fact]
        public void AssemblySuggestions_IncludesIdentityAndRouting()
        {
            var recon = new ReconciliationResult();
            recon.GapFills.Add(new GapFill { Field = "Description", Value = "WELDMENT ASM", Source = "PDF", Confidence = 0.85 });
            recon.RoutingSuggestions.Add(new RoutingSuggestion
            {
                Operation = RoutingOp.Weld,
                WorkCenter = "F400",
                NoteText = "WELD PER DWG",
                SourceNote = "WELD PER DWG",
                Confidence = 0.90
            });

            var suggestions = _service.GenerateAssemblySuggestions(recon, new Dictionary<string, string>());

            // Identity
            Assert.Contains(suggestions, s => s.PropertyName == "Description" && s.Value == "WELDMENT ASM");

            // Routing
            Assert.Contains(suggestions, s => s.PropertyName == "OP20_WC" && s.Value == "F400");
            Assert.Contains(suggestions, s => s.PropertyName == "OP20_RN" && s.Value == "WELD PER DWG");
        }

        // --- Edge cases ---

        [Fact]
        public void EmptyReconciliation_ReturnsEmptySuggestions()
        {
            var recon = new ReconciliationResult();
            var suggestions = _service.GeneratePartSuggestions(recon, new Dictionary<string, string>());
            Assert.Empty(suggestions);
        }

        [Fact]
        public void NullCurrentProperties_HandledGracefully()
        {
            var recon = new ReconciliationResult();
            recon.GapFills.Add(new GapFill { Field = "Description", Value = "TEST", Source = "PDF", Confidence = 0.85 });

            var suggestions = _service.GeneratePartSuggestions(recon, null);
            Assert.Single(suggestions);
        }

        [Fact]
        public void PropertySuggestion_IsGapFill_WhenCurrentEmpty()
        {
            var s = new PropertySuggestion { Value = "NEW", CurrentValue = null };
            Assert.True(s.IsGapFill);
            Assert.False(s.IsOverride);
        }

        [Fact]
        public void PropertySuggestion_IsOverride_WhenCurrentDiffers()
        {
            var s = new PropertySuggestion { Value = "NEW", CurrentValue = "OLD" };
            Assert.False(s.IsGapFill);
            Assert.True(s.IsOverride);
        }

        [Fact]
        public void PropertySuggestion_NotOverride_WhenSameValue()
        {
            var s = new PropertySuggestion { Value = "SAME", CurrentValue = "SAME" };
            Assert.False(s.IsGapFill);
            Assert.False(s.IsOverride);
        }
    }
}
