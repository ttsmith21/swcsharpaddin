using System.Linq;
using Xunit;
using NM.Core.DataModel;
using NM.Core.Pdf.Models;
using NM.Core.Reconciliation;
using NM.Core.Reconciliation.Models;

namespace NM.Core.Tests.Reconciliation
{
    public class ReconciliationEngineTests
    {
        private readonly ReconciliationEngine _engine = new ReconciliationEngine();

        // --- Material comparison ---

        [Fact]
        public void MatchingMaterial_AddConfirmation()
        {
            var model = new PartData { Material = "304 SS" };
            var drawing = new DrawingData { Material = "304 SS" };

            var result = _engine.Reconcile(model, drawing);

            Assert.Contains(result.Confirmations, c => c.Contains("Material"));
            Assert.DoesNotContain(result.Conflicts, c => c.Field == "Material");
        }

        [Fact]
        public void EquivalentMaterial_AddConfirmation()
        {
            var model = new PartData { Material = "304" };
            var drawing = new DrawingData { Material = "304 STAINLESS STEEL" };

            var result = _engine.Reconcile(model, drawing);

            Assert.Contains(result.Confirmations, c => c.Contains("Material"));
        }

        [Fact]
        public void DifferentMaterial_AddConflict()
        {
            var model = new PartData { Material = "304 SS" };
            var drawing = new DrawingData { Material = "A36 CARBON STEEL" };

            var result = _engine.Reconcile(model, drawing);

            Assert.True(result.HasConflicts);
            var conflict = result.Conflicts.First(c => c.Field == "Material");
            Assert.Equal("304 SS", conflict.ModelValue);
            Assert.Equal("A36 CARBON STEEL", conflict.DrawingValue);
            Assert.Equal(ConflictSeverity.High, conflict.Severity);
        }

        // --- Thickness comparison ---

        [Fact]
        public void MatchingThickness_AddConfirmation()
        {
            var model = new PartData { Thickness_m = 0.003175 }; // 0.125"
            var drawing = new DrawingData { Thickness_in = 0.125 };

            var result = _engine.Reconcile(model, drawing);

            Assert.Contains(result.Confirmations, c => c.Contains("Thickness"));
        }

        [Fact]
        public void DifferentThickness_AddConflict()
        {
            var model = new PartData { Thickness_m = 0.003175 }; // 0.125"
            var drawing = new DrawingData { Thickness_in = 0.250 };

            var result = _engine.Reconcile(model, drawing);

            var conflict = result.Conflicts.First(c => c.Field == "Thickness");
            Assert.Equal(ConflictSeverity.High, conflict.Severity);
            Assert.Equal(ConflictResolution.UseModel, conflict.Recommendation);
        }

        // --- Gap filling ---

        [Fact]
        public void MissingDescription_GapFilled()
        {
            var model = new PartData(); // No description in Extra
            var drawing = new DrawingData { Description = "MOUNTING BRACKET" };

            var result = _engine.Reconcile(model, drawing);

            Assert.True(result.HasGapFills);
            var gap = result.GapFills.First(g => g.Field == "Description");
            Assert.Equal("MOUNTING BRACKET", gap.Value);
            Assert.Equal("PDF title block", gap.Source);
        }

        [Fact]
        public void MissingRevision_GapFilled()
        {
            var model = new PartData();
            var drawing = new DrawingData { Revision = "C" };

            var result = _engine.Reconcile(model, drawing);

            var gap = result.GapFills.First(g => g.Field == "Revision");
            Assert.Equal("C", gap.Value);
            Assert.True(gap.Confidence >= 0.85);
        }

        [Fact]
        public void MissingPartNumber_GapFilled()
        {
            var model = new PartData();
            var drawing = new DrawingData { PartNumber = "NM-1234" };

            var result = _engine.Reconcile(model, drawing);

            var gap = result.GapFills.First(g => g.Field == "PartNumber");
            Assert.Equal("NM-1234", gap.Value);
        }

        [Fact]
        public void ExistingField_NotGapFilled()
        {
            var model = new PartData { Material = "304 SS" };
            var drawing = new DrawingData { Material = "316 SS" };

            var result = _engine.Reconcile(model, drawing);

            // Material exists in model — should be a conflict, not a gap fill
            Assert.DoesNotContain(result.GapFills, g => g.Field == "Material");
        }

        // --- Routing suggestions ---

        [Fact]
        public void DrawingNotes_GenerateRoutingSuggestions()
        {
            var model = new PartData();
            var drawing = new DrawingData();
            drawing.RoutingHints.Add(new RoutingHint
            {
                Operation = RoutingOp.Deburr,
                WorkCenter = "F210",
                NoteText = "BREAK ALL EDGES",
                SourceNote = "BREAK ALL EDGES",
                Confidence = 0.95
            });

            var result = _engine.Reconcile(model, drawing);

            Assert.True(result.HasRoutingSuggestions);
            var suggestion = result.RoutingSuggestions[0];
            Assert.Equal(RoutingOp.Deburr, suggestion.Operation);
            Assert.Equal("F210", suggestion.WorkCenter);
            Assert.Equal(30, suggestion.SuggestedOpNumber); // OP30 for deburr
        }

        [Fact]
        public void MultipleRoutingHints_SortedByOpNumber()
        {
            var drawing = new DrawingData();
            drawing.RoutingHints.Add(new RoutingHint { Operation = RoutingOp.Weld, WorkCenter = "F400", NoteText = "WELD", Confidence = 0.9 });
            drawing.RoutingHints.Add(new RoutingHint { Operation = RoutingOp.Deburr, WorkCenter = "F210", NoteText = "DEBURR", Confidence = 0.9 });
            drawing.RoutingHints.Add(new RoutingHint { Operation = RoutingOp.OutsideProcess, NoteText = "GALVANIZE", Confidence = 0.9 });

            var result = _engine.Reconcile(new PartData(), drawing);

            Assert.Equal(3, result.RoutingSuggestions.Count);
            // Should be sorted: Deburr(30), Weld(50), OutsideProcess(60)
            Assert.Equal(30, result.RoutingSuggestions[0].SuggestedOpNumber);
            Assert.Equal(50, result.RoutingSuggestions[1].SuggestedOpNumber);
            Assert.Equal(60, result.RoutingSuggestions[2].SuggestedOpNumber);
        }

        // --- Rename suggestion ---

        [Fact]
        public void DifferentFilename_SuggestsRename()
        {
            var model = new PartData { FilePath = @"C:\Parts\Part1.sldprt" };
            var drawing = new DrawingData
            {
                PartNumber = "NM-1234",
                Description = "BRACKET"
            };

            var result = _engine.Reconcile(model, drawing);

            Assert.True(result.HasRenameSuggestion);
            Assert.Contains("NM-1234", result.Rename.NewPath);
            Assert.Contains("BRACKET", result.Rename.NewPath);
            Assert.True(result.Rename.RequiresUserApproval);
        }

        [Fact]
        public void MatchingFilename_NoRenameSuggested()
        {
            var model = new PartData { FilePath = @"C:\Parts\NM-1234.sldprt" };
            var drawing = new DrawingData { PartNumber = "NM-1234" };

            var result = _engine.Reconcile(model, drawing);

            Assert.False(result.HasRenameSuggestion);
        }

        // --- Edge cases ---

        [Fact]
        public void NullDrawing_ReturnsEmptyResult()
        {
            var result = _engine.Reconcile(new PartData(), null);

            Assert.False(result.HasActions);
            Assert.Empty(result.Conflicts);
            Assert.Empty(result.GapFills);
        }

        [Fact]
        public void NullModel_StillExtractsDrawingData()
        {
            var drawing = new DrawingData
            {
                PartNumber = "12345",
                Description = "PLATE",
                Material = "A36"
            };

            var result = _engine.Reconcile(null, drawing);

            Assert.True(result.HasGapFills);
            Assert.True(result.GapFills.Count >= 3);
        }

        [Fact]
        public void CompleteReconciliation_Summary()
        {
            var model = new PartData
            {
                Material = "304 SS",
                Thickness_m = 0.003175, // 0.125"
                FilePath = @"C:\Parts\Part1.sldprt"
            };
            var drawing = new DrawingData
            {
                PartNumber = "NM-5678",
                Description = "MOUNTING BRACKET",
                Revision = "B",
                Material = "304 STAINLESS STEEL",
                Thickness_in = 0.125
            };
            drawing.RoutingHints.Add(new RoutingHint
            {
                Operation = RoutingOp.Deburr,
                WorkCenter = "F210",
                NoteText = "BREAK ALL EDGES",
                Confidence = 0.95
            });

            var result = _engine.Reconcile(model, drawing);

            // Should have: confirmations (material, thickness), gap fills (PN, desc, rev),
            // routing suggestion (deburr), rename suggestion
            Assert.NotEmpty(result.Confirmations);
            Assert.True(result.HasGapFills);
            Assert.True(result.HasRoutingSuggestions);
            Assert.True(result.HasRenameSuggestion);
            Assert.NotEmpty(result.Summary);
        }
    }
}
