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
        public void PartSuggestions_MapsPartNumberToPrint()
        {
            var recon = new ReconciliationResult();
            recon.GapFills.Add(new GapFill { Field = "PartNumber", Value = "12345-01", Source = "PDF title block", Confidence = 0.90 });

            var suggestions = _service.GeneratePartSuggestions(recon, new Dictionary<string, string>());

            // PartNumber maps to "Print" property name (tab builder field)
            Assert.Contains(suggestions, s => s.PropertyName == "Print" && s.Value == "12345-01");
        }

        [Fact]
        public void PartSuggestions_MapsMaterialToOptiMaterial()
        {
            var recon = new ReconciliationResult();
            recon.GapFills.Add(new GapFill { Field = "Material", Value = "304 SS", Source = "PDF title block", Confidence = 0.85 });

            var suggestions = _service.GeneratePartSuggestions(recon, new Dictionary<string, string>());

            // Material maps to "OptiMaterial" (NOT rbMaterialType which is a RadioButton 0/1/2)
            Assert.Contains(suggestions, s => s.PropertyName == "OptiMaterial" && s.Value == "304 SS");
            Assert.DoesNotContain(suggestions, s => s.PropertyName == "rbMaterialType");
        }

        // --- Part routing: fixed slots ---

        [Fact]
        public void PartSuggestions_DeburNote_EnablesF210AndSetsRN()
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

            // Should enable F210 checkbox
            Assert.Contains(suggestions, s => s.PropertyName == "F210" && s.Value == "1");
            // Should set routing note to F210_RN (not F210_Note)
            var rn = suggestions.First(s => s.PropertyName == "F210_RN");
            Assert.Equal("BREAK ALL EDGES", rn.Value);
            Assert.Equal(PropertyCategory.Routing, rn.Category);
        }

        [Fact]
        public void PartSuggestions_TapNote_EnablesF220AndSetsRN()
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

            // Should enable F220 checkbox
            Assert.Contains(suggestions, s => s.PropertyName == "F220" && s.Value == "1");
            // Should set routing note to F220_RN (not F220_Note)
            Assert.Contains(suggestions, s => s.PropertyName == "F220_RN" && s.Value.Contains("TAP 1/4-20"));
        }

        [Fact]
        public void PartSuggestions_ProcessOverride_MapsToOP20()
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

            // Work center goes to "OP20" (the ComboBox), not "OP20_WorkCenter"
            Assert.Contains(suggestions, s => s.PropertyName == "OP20" && s.Value == "F110");
            // Routing note goes to OP20_RN
            Assert.Contains(suggestions, s => s.PropertyName == "OP20_RN" && s.Value == "WATERJET ONLY");
        }

        // --- Part routing: outsource ---

        [Fact]
        public void PartSuggestions_OutsideProcess_MapsToOS()
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

            // Outside process routes to OS_RN
            Assert.Contains(suggestions, s => s.PropertyName == "OS_RN" && s.Value == "GALVANIZE");
        }

        [Fact]
        public void PartSuggestions_FinishNote_MapsToOSWithWorkCenter()
        {
            var recon = new ReconciliationResult();
            recon.RoutingSuggestions.Add(new RoutingSuggestion
            {
                Operation = RoutingOp.Finish,
                WorkCenter = "N140",
                NoteText = "POWDER COAT BLACK",
                SourceNote = "POWDER COAT BLACK",
                Confidence = 0.85
            });

            var suggestions = _service.GeneratePartSuggestions(recon, new Dictionary<string, string>());

            Assert.Contains(suggestions, s => s.PropertyName == "OS_WC" && s.Value == "N140");
            Assert.Contains(suggestions, s => s.PropertyName == "OS_RN" && s.Value == "POWDER COAT BLACK");
        }

        // --- Part routing: Other WC slots ---

        [Fact]
        public void PartSuggestions_WeldNote_UsesOtherWCSlot1()
        {
            var recon = new ReconciliationResult();
            recon.RoutingSuggestions.Add(new RoutingSuggestion
            {
                Operation = RoutingOp.Weld,
                WorkCenter = "F400",
                NoteText = "WELD PER DWG",
                SourceNote = "WELD PER DWG",
                Confidence = 0.90
            });

            var suggestions = _service.GeneratePartSuggestions(recon, new Dictionary<string, string>());

            // Slot 1 uses: OtherWC_CB, OtherOP, Other_WC, Other_RN
            Assert.Contains(suggestions, s => s.PropertyName == "OtherWC_CB" && s.Value == "1");
            Assert.Contains(suggestions, s => s.PropertyName == "OtherOP" && s.Value == "60");
            Assert.Contains(suggestions, s => s.PropertyName == "Other_WC" && s.Value == "F400");
            Assert.Contains(suggestions, s => s.PropertyName == "Other_RN" && s.Value == "WELD PER DWG");
        }

        [Fact]
        public void PartSuggestions_TwoOtherOps_UsesSlots1And2()
        {
            var recon = new ReconciliationResult();
            recon.RoutingSuggestions.Add(new RoutingSuggestion
            {
                Operation = RoutingOp.Weld,
                WorkCenter = "F400",
                NoteText = "WELD PER DWG",
                SourceNote = "WELD PER DWG",
                Confidence = 0.90
            });
            recon.RoutingSuggestions.Add(new RoutingSuggestion
            {
                Operation = RoutingOp.Inspect,
                NoteText = "INSPECT ALL WELDS",
                SourceNote = "INSPECT ALL WELDS",
                Confidence = 0.85
            });

            var suggestions = _service.GeneratePartSuggestions(recon, new Dictionary<string, string>());

            // Slot 1 (Weld)
            Assert.Contains(suggestions, s => s.PropertyName == "OtherWC_CB" && s.Value == "1");
            Assert.Contains(suggestions, s => s.PropertyName == "Other_WC" && s.Value == "F400");
            Assert.Contains(suggestions, s => s.PropertyName == "Other_RN" && s.Value == "WELD PER DWG");

            // Slot 2 (Inspect) — uses {N} suffix pattern
            Assert.Contains(suggestions, s => s.PropertyName == "OtherWC_CB2" && s.Value == "1");
            Assert.Contains(suggestions, s => s.PropertyName == "Other_OP2" && s.Value == "70");
            Assert.Contains(suggestions, s => s.PropertyName == "Other_RN2" && s.Value == "INSPECT ALL WELDS");
        }

        // --- Appending notes ---

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
                { "F210_RN", "BREAK ALL EDGES" }
            };

            var suggestions = _service.GeneratePartSuggestions(recon, currentProps);

            var note = suggestions.First(s => s.PropertyName == "F210_RN");
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
        public void AssemblyOperations_ToProperties_UsesOP_NotOP_WC()
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

            // Work center is stored in OP## (e.g., "OP30"), NOT "OP30_WC"
            Assert.Equal("F400", props["OP30"]);
            Assert.False(props.ContainsKey("OP30_WC"));
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

            // Routing — uses OP## for work center (not OP##_WC)
            Assert.Contains(suggestions, s => s.PropertyName == "OP20" && s.Value == "F400");
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

        [Fact]
        public void PartSuggestions_OtherWCSlot_IncludesSetupAndRun()
        {
            var recon = new ReconciliationResult();
            recon.RoutingSuggestions.Add(new RoutingSuggestion
            {
                Operation = RoutingOp.Weld,
                WorkCenter = "F400",
                NoteText = "WELD PER DWG",
                SourceNote = "WELD PER DWG",
                Confidence = 0.90
            });

            var suggestions = _service.GeneratePartSuggestions(recon, new Dictionary<string, string>());

            // Weld default: setup=15, run=10
            Assert.Contains(suggestions, s => s.PropertyName == "Other_S" && s.Value == "15");
            Assert.Contains(suggestions, s => s.PropertyName == "Other_R" && s.Value == "10");
        }
    }
}
