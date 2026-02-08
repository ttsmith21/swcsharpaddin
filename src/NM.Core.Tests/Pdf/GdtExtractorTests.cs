using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NM.Core.Pdf;
using NM.Core.Pdf.Models;

namespace NM.Core.Tests.Pdf
{
    [TestClass]
    public class GdtExtractorTests
    {
        private readonly GdtExtractor _extractor = new GdtExtractor();

        // =====================================================================
        // True Position extraction
        // =====================================================================

        [TestMethod]
        public void Extract_TruePosition_ParsesTolerance()
        {
            var callouts = _extractor.Extract("TRUE POSITION .005 A B C");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.Position, callouts[0].GdtFeatureType);
            Assert.AreEqual(0.005, callouts[0].ToleranceValue.Value, 0.0001);
            Assert.AreEqual("position", callouts[0].Type);
        }

        [TestMethod]
        public void Extract_TruePosition_WithDatums()
        {
            var callouts = _extractor.Extract("TRUE POSITION .005 A B C");
            Assert.IsTrue(callouts[0].DatumReferences.Contains("A"));
            Assert.IsTrue(callouts[0].DatumReferences.Contains("B"));
            Assert.IsTrue(callouts[0].DatumReferences.Contains("C"));
        }

        [TestMethod]
        public void Extract_TruePosition_WithMmc()
        {
            var callouts = _extractor.Extract("TRUE POSITION .005 MMC A B");
            Assert.AreEqual(1, callouts.Count);
            Assert.IsTrue(callouts[0].IsMmc);
            Assert.IsFalse(callouts[0].IsLmc);
        }

        [TestMethod]
        public void Extract_TruePosition_WithLmc()
        {
            var callouts = _extractor.Extract("TRUE POSITION .003 LMC A");
            Assert.AreEqual(1, callouts.Count);
            Assert.IsTrue(callouts[0].IsLmc);
        }

        [TestMethod]
        public void Extract_TruePosition_TightClassification()
        {
            var callouts = _extractor.Extract("TRUE POSITION .003 A B");
            Assert.AreEqual(ToleranceTier.Precision, callouts[0].Tier);
            Assert.AreEqual(CostImpact.Critical, callouts[0].Impact);
        }

        [TestMethod]
        public void Extract_TruePosition_ModerateClassification()
        {
            var callouts = _extractor.Extract("TRUE POSITION .010 A B");
            Assert.AreEqual(ToleranceTier.Moderate, callouts[0].Tier);
        }

        [TestMethod]
        public void Extract_TruePosition_StandardClassification()
        {
            var callouts = _extractor.Extract("TRUE POSITION .020 A B");
            Assert.AreEqual(ToleranceTier.Standard, callouts[0].Tier);
        }

        [TestMethod]
        public void Extract_TruePosition_MmcRelaxesTier()
        {
            // 0.005 at MMC: effective = 0.005 * 1.5 = 0.0075
            // 0.0075 > 0.007 → Moderate (not Tight)
            var callouts = _extractor.Extract("TRUE POSITION .005 MMC A B");
            Assert.AreEqual(ToleranceTier.Moderate, callouts[0].Tier);
        }

        [TestMethod]
        public void Extract_TruePosition_Abbreviation()
        {
            var callouts = _extractor.Extract("T/P .005 A B");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.Position, callouts[0].GdtFeatureType);
        }

        // =====================================================================
        // Flatness extraction
        // =====================================================================

        [TestMethod]
        public void Extract_Flatness_ParsesTolerance()
        {
            var callouts = _extractor.Extract("FLATNESS .002");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.Flatness, callouts[0].GdtFeatureType);
            Assert.AreEqual(0.002, callouts[0].ToleranceValue.Value, 0.0001);
        }

        [TestMethod]
        public void Extract_Flatness_TightIsHighImpact()
        {
            var callouts = _extractor.Extract("FLATNESS .001");
            Assert.AreEqual(ToleranceTier.Precision, callouts[0].Tier);
            Assert.AreEqual(CostImpact.Critical, callouts[0].Impact);
        }

        [TestMethod]
        public void Extract_Flatness_StandardIsLow()
        {
            var callouts = _extractor.Extract("FLATNESS .010");
            Assert.AreEqual(ToleranceTier.Standard, callouts[0].Tier);
            Assert.AreEqual(CostImpact.Low, callouts[0].Impact);
        }

        // =====================================================================
        // Perpendicularity extraction
        // =====================================================================

        [TestMethod]
        public void Extract_Perpendicularity_ParsesTolerance()
        {
            var callouts = _extractor.Extract("PERPENDICULARITY .003 A");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.Perpendicularity, callouts[0].GdtFeatureType);
            Assert.AreEqual(0.003, callouts[0].ToleranceValue.Value, 0.0001);
        }

        [TestMethod]
        public void Extract_Perpendicularity_WithDatum()
        {
            var callouts = _extractor.Extract("PERPENDICULARITY .003 A");
            Assert.AreEqual(1, callouts[0].DatumReferences.Count);
            Assert.AreEqual("A", callouts[0].DatumReferences[0]);
        }

        // =====================================================================
        // Parallelism extraction
        // =====================================================================

        [TestMethod]
        public void Extract_Parallelism_ParsesTolerance()
        {
            var callouts = _extractor.Extract("PARALLELISM .005 A");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.Parallelism, callouts[0].GdtFeatureType);
            Assert.AreEqual(ToleranceTier.Tight, callouts[0].Tier);
        }

        // =====================================================================
        // Concentricity / Runout
        // =====================================================================

        [TestMethod]
        public void Extract_Concentricity_ParsesTolerance()
        {
            var callouts = _extractor.Extract("CONCENTRICITY .003 A");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.Concentricity, callouts[0].GdtFeatureType);
        }

        [TestMethod]
        public void Extract_TotalRunout_ParsesTolerance()
        {
            var callouts = _extractor.Extract("TOTAL RUNOUT .002 A");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.TotalRunout, callouts[0].GdtFeatureType);
            Assert.AreEqual(ToleranceTier.Precision, callouts[0].Tier);
        }

        [TestMethod]
        public void Extract_CircularRunout_ParsesTolerance()
        {
            var callouts = _extractor.Extract("RUNOUT .005 A");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.CircularRunout, callouts[0].GdtFeatureType);
        }

        // =====================================================================
        // Profile extraction
        // =====================================================================

        [TestMethod]
        public void Extract_ProfileOfSurface_ParsesTolerance()
        {
            var callouts = _extractor.Extract("PROFILE OF A SURFACE .003 A B");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.ProfileOfSurface, callouts[0].GdtFeatureType);
        }

        [TestMethod]
        public void Extract_ProfileOfLine_ParsesTolerance()
        {
            var callouts = _extractor.Extract("PROFILE OF A LINE .005 A");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.ProfileOfLine, callouts[0].GdtFeatureType);
        }

        [TestMethod]
        public void Extract_ProfileGeneric_DefaultsToSurface()
        {
            var callouts = _extractor.Extract("PROFILE .004 A");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.ProfileOfSurface, callouts[0].GdtFeatureType);
        }

        // =====================================================================
        // Other form tolerances
        // =====================================================================

        [TestMethod]
        public void Extract_Straightness_ParsesTolerance()
        {
            var callouts = _extractor.Extract("STRAIGHTNESS .002");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.Straightness, callouts[0].GdtFeatureType);
        }

        [TestMethod]
        public void Extract_Circularity_ParsesTolerance()
        {
            var callouts = _extractor.Extract("CIRCULARITY .003");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.Circularity, callouts[0].GdtFeatureType);
        }

        [TestMethod]
        public void Extract_Roundness_MapsToCircularity()
        {
            var callouts = _extractor.Extract("ROUNDNESS .003");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.Circularity, callouts[0].GdtFeatureType);
        }

        [TestMethod]
        public void Extract_Cylindricity_ParsesTolerance()
        {
            var callouts = _extractor.Extract("CYLINDRICITY .004");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.Cylindricity, callouts[0].GdtFeatureType);
        }

        [TestMethod]
        public void Extract_Angularity_ParsesTolerance()
        {
            var callouts = _extractor.Extract("ANGULARITY .005 A");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.Angularity, callouts[0].GdtFeatureType);
        }

        // =====================================================================
        // Cost flags
        // =====================================================================

        [TestMethod]
        public void ToCostFlags_TightPosition_GeneratesFlag()
        {
            var callouts = _extractor.Extract("TRUE POSITION .003 A B C");
            var flags = _extractor.ToCostFlags(callouts);
            Assert.IsTrue(flags.Count > 0);
            Assert.IsTrue(flags.Any(f => f.Source == "GD&T callout"));
            Assert.IsTrue(flags.Any(f => f.Impact >= CostImpact.High));
        }

        [TestMethod]
        public void ToCostFlags_StandardTolerance_NoFlag()
        {
            var callouts = _extractor.Extract("FLATNESS .010");
            var flags = _extractor.ToCostFlags(callouts);
            Assert.AreEqual(0, flags.Count); // Standard tier = no flag
        }

        [TestMethod]
        public void ToCostFlags_ModerateGeneratesFlag()
        {
            var callouts = _extractor.Extract("FLATNESS .005");
            var flags = _extractor.ToCostFlags(callouts);
            Assert.AreEqual(1, flags.Count); // Moderate tier ≥ Moderate → flag
        }

        // =====================================================================
        // Routing hints
        // =====================================================================

        [TestMethod]
        public void ToRoutingHints_TightPosition_AddsCmmAndFixture()
        {
            var callouts = _extractor.Extract("TRUE POSITION .003 A B");
            var hints = _extractor.ToRoutingHints(callouts);
            Assert.IsTrue(hints.Any(h => h.NoteText.Contains("CMM")));
            Assert.IsTrue(hints.Any(h => h.NoteText.Contains("FIXTURE")));
        }

        [TestMethod]
        public void ToRoutingHints_StandardTolerance_NoHints()
        {
            var callouts = _extractor.Extract("FLATNESS .010");
            var hints = _extractor.ToRoutingHints(callouts);
            Assert.AreEqual(0, hints.Count);
        }

        [TestMethod]
        public void ToRoutingHints_MultipleTight_FlagsReview()
        {
            string text = "TRUE POSITION .003 A B\nFLATNESS .001\nPERPENDICULARITY .002 A";
            var callouts = _extractor.Extract(text);
            var hints = _extractor.ToRoutingHints(callouts);
            Assert.IsTrue(hints.Any(h => h.NoteText.Contains("MULTIPLE TIGHT GD&T")));
        }

        // =====================================================================
        // Tier classification static methods
        // =====================================================================

        [TestMethod]
        public void ClassifyPositionTier_VeryTight()
        {
            Assert.AreEqual(ToleranceTier.Precision, GdtExtractor.ClassifyPositionTier(0.002, false));
        }

        [TestMethod]
        public void ClassifyPositionTier_Tight()
        {
            Assert.AreEqual(ToleranceTier.Tight, GdtExtractor.ClassifyPositionTier(0.005, false));
        }

        [TestMethod]
        public void ClassifyPositionTier_MmcRelaxes()
        {
            // 0.005 with MMC: effective = 0.0075 → Moderate
            Assert.AreEqual(ToleranceTier.Moderate, GdtExtractor.ClassifyPositionTier(0.005, true));
        }

        [TestMethod]
        public void ClassifyFormTier_Precision()
        {
            Assert.AreEqual(ToleranceTier.Precision, GdtExtractor.ClassifyFormTier(0.001));
        }

        [TestMethod]
        public void ClassifyFormTier_Standard()
        {
            Assert.AreEqual(ToleranceTier.Standard, GdtExtractor.ClassifyFormTier(0.010));
        }

        [TestMethod]
        public void ClassifyOrientationTier_Tight()
        {
            Assert.AreEqual(ToleranceTier.Tight, GdtExtractor.ClassifyOrientationTier(0.005));
        }

        [TestMethod]
        public void ClassifyRunoutTier_Precision()
        {
            Assert.AreEqual(ToleranceTier.Precision, GdtExtractor.ClassifyRunoutTier(0.001));
        }

        // =====================================================================
        // Edge cases
        // =====================================================================

        [TestMethod]
        public void Extract_NullText_ReturnsEmpty()
        {
            Assert.AreEqual(0, _extractor.Extract(null).Count);
        }

        [TestMethod]
        public void Extract_EmptyText_ReturnsEmpty()
        {
            Assert.AreEqual(0, _extractor.Extract("").Count);
        }

        [TestMethod]
        public void Extract_NoGdt_ReturnsEmpty()
        {
            Assert.AreEqual(0, _extractor.Extract("This is a regular drawing with no GD&T").Count);
        }

        [TestMethod]
        public void Extract_Deduplicates_SameTypeAndValue()
        {
            // Same callout appearing twice should be deduped
            string text = "TRUE POSITION .005 A B\nSome text\nTRUE POSITION .005 A B";
            var callouts = _extractor.Extract(text);
            Assert.AreEqual(1, callouts.Count);
        }

        [TestMethod]
        public void Extract_DifferentValues_NotDeduped()
        {
            string text = "TRUE POSITION .005 A B\nTRUE POSITION .003 A B";
            var callouts = _extractor.Extract(text);
            Assert.AreEqual(2, callouts.Count);
        }

        [TestMethod]
        public void Extract_SortsTightestFirst()
        {
            string text = "FLATNESS .010\nFLATNESS .001\nFLATNESS .005";
            var callouts = _extractor.Extract(text);
            Assert.AreEqual(3, callouts.Count);
            Assert.AreEqual(0.001, callouts[0].ToleranceValue.Value, 0.0001);
            Assert.AreEqual(0.005, callouts[1].ToleranceValue.Value, 0.0001);
            Assert.AreEqual(0.010, callouts[2].ToleranceValue.Value, 0.0001);
        }

        // =====================================================================
        // Realistic drawing scenario
        // =====================================================================

        [TestMethod]
        public void Extract_RealisticDrawing_MultipleCallouts()
        {
            string drawingText = @"
                NOTES:
                1. BREAK ALL SHARP EDGES .005/.015
                2. PART TO BE ANODIZED PER MIL-A-8625 TYPE III
                3. ALL MACHINED SURFACES 63 Ra UNLESS OTHERWISE NOTED

                TRUE POSITION .005 MMC A B C
                FLATNESS .002
                PERPENDICULARITY .003 A
                PROFILE OF A SURFACE .005 A B

                UNLESS OTHERWISE SPECIFIED:
                .XX ±0.01
                .XXX ±0.005
                ANGLES ±1°
            ";

            var callouts = _extractor.Extract(drawingText);

            // Should find: position, flatness, perpendicularity, profile
            Assert.IsTrue(callouts.Count >= 4,
                $"Expected at least 4 GD&T callouts, got {callouts.Count}");

            Assert.IsTrue(callouts.Any(c => c.GdtFeatureType == GdtType.Position));
            Assert.IsTrue(callouts.Any(c => c.GdtFeatureType == GdtType.Flatness));
            Assert.IsTrue(callouts.Any(c => c.GdtFeatureType == GdtType.Perpendicularity));
            Assert.IsTrue(callouts.Any(c => c.GdtFeatureType == GdtType.ProfileOfSurface));

            // Cost flags should be generated for tight ones
            var flags = _extractor.ToCostFlags(callouts);
            Assert.IsTrue(flags.Count > 0, "Expected cost flags for tight GD&T callouts");

            // Routing hints should include CMM
            var hints = _extractor.ToRoutingHints(callouts);
            Assert.IsTrue(hints.Any(h => h.NoteText.Contains("CMM")),
                "Expected CMM routing hint for tight GD&T");
        }

        [TestMethod]
        public void Extract_UnicodeSymbols_Position()
        {
            // PDF text extraction sometimes preserves Unicode GD&T symbols
            var callouts = _extractor.Extract("⌖ .005 A B C");
            Assert.AreEqual(1, callouts.Count);
            Assert.AreEqual(GdtType.Position, callouts[0].GdtFeatureType);
        }

        [TestMethod]
        public void Extract_WithDiameterSymbol()
        {
            var callouts = _extractor.Extract("TRUE POSITION ⌀.005 A B");
            Assert.AreEqual(1, callouts.Count);
            Assert.IsTrue(callouts[0].IsDiametral);
        }
    }
}
