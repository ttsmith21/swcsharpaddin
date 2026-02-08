using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NM.Core.Pdf;
using NM.Core.Pdf.Models;
using static NM.Core.Pdf.IsoToleranceStandard;

namespace NM.Core.Tests.Pdf
{
    [TestClass]
    public class IsoToleranceStandardTests
    {
        // =====================================================================
        // ISO 13920 Linear Tolerance Table
        // =====================================================================

        [TestMethod]
        public void Iso13920Linear_SmallDim_ClassB()
        {
            // 25mm dimension, Class B → ±2mm
            double tol = GetIso13920Linear(25, Iso13920Linear.B);
            Assert.AreEqual(2, tol);
        }

        [TestMethod]
        public void Iso13920Linear_SmallDim_ClassA()
        {
            // 25mm dimension, Class A → ±1mm
            double tol = GetIso13920Linear(25, Iso13920Linear.A);
            Assert.AreEqual(1, tol);
        }

        [TestMethod]
        public void Iso13920Linear_MediumDim_ClassB()
        {
            // 500mm dimension, Class B → ±4mm
            double tol = GetIso13920Linear(500, Iso13920Linear.B);
            Assert.AreEqual(4, tol);
        }

        [TestMethod]
        public void Iso13920Linear_LargeDim_ClassC()
        {
            // 1500mm dimension, Class C → ±11mm
            double tol = GetIso13920Linear(1500, Iso13920Linear.C);
            Assert.AreEqual(11, tol);
        }

        [TestMethod]
        public void Iso13920Linear_ClassD_Loosest()
        {
            // 500mm dimension, Class D → ±12mm
            double tol = GetIso13920Linear(500, Iso13920Linear.D);
            Assert.AreEqual(12, tol);
        }

        // =====================================================================
        // ISO 13920 Geometric Tolerance Table
        // =====================================================================

        [TestMethod]
        public void Iso13920Geometric_ClassF_Medium()
        {
            // 500mm span, Class F → 3mm
            double tol = GetIso13920Geometric(500, Iso13920Geometric.F);
            Assert.AreEqual(3, tol);
        }

        [TestMethod]
        public void Iso13920Geometric_ClassE_Fine()
        {
            // 500mm span, Class E → 1.5mm
            double tol = GetIso13920Geometric(500, Iso13920Geometric.E);
            Assert.AreEqual(1.5, tol);
        }

        [TestMethod]
        public void Iso13920Geometric_ClassH_VeryCoarse()
        {
            // 100mm span, Class H → 2.5mm
            double tol = GetIso13920Geometric(100, Iso13920Geometric.H);
            Assert.AreEqual(2.5, tol);
        }

        // =====================================================================
        // ISO 2768 Linear Tolerance Table
        // =====================================================================

        [TestMethod]
        public void Iso2768Linear_Medium_50mm()
        {
            // 50mm, class m → ±0.3mm
            double tol = GetIso2768Linear(50, Iso2768Linear.Medium);
            Assert.AreEqual(0.3, tol, 0.001);
        }

        [TestMethod]
        public void Iso2768Linear_Fine_200mm()
        {
            // 200mm, class f → ±0.2mm
            double tol = GetIso2768Linear(200, Iso2768Linear.Fine);
            Assert.AreEqual(0.2, tol, 0.001);
        }

        [TestMethod]
        public void Iso2768Linear_Coarse_10mm()
        {
            // 10mm, class c → ±0.5mm
            double tol = GetIso2768Linear(10, Iso2768Linear.Coarse);
            Assert.AreEqual(0.5, tol, 0.001);
        }

        // =====================================================================
        // Class comparison
        // =====================================================================

        [TestMethod]
        public void IsTighterLinear_A_TighterThan_B()
        {
            Assert.IsTrue(IsTighterLinear(Iso13920Linear.A, Iso13920Linear.B));
        }

        [TestMethod]
        public void IsTighterLinear_B_NotTighterThan_A()
        {
            Assert.IsFalse(IsTighterLinear(Iso13920Linear.B, Iso13920Linear.A));
        }

        [TestMethod]
        public void IsTighterGeometric_E_TighterThan_F()
        {
            Assert.IsTrue(IsTighterGeometric(Iso13920Geometric.E, Iso13920Geometric.F));
        }

        // =====================================================================
        // Classification from tolerance value
        // =====================================================================

        [TestMethod]
        public void ClassifyLinear_TightTolerance_ReturnsClassA()
        {
            // 500mm nominal, ±1.5mm band → ±0.75mm → fits Class A (±2mm)
            var cls = ClassifyLinearTolerance13920(500, 3.0);
            Assert.AreEqual(Iso13920Linear.A, cls);
        }

        [TestMethod]
        public void ClassifyLinear_StandardTolerance_ReturnsClassB()
        {
            // 500mm nominal, ±4mm band → fits Class B
            var cls = ClassifyLinearTolerance13920(500, 8.0);
            Assert.AreEqual(Iso13920Linear.B, cls);
        }

        [TestMethod]
        public void ClassifyLinear_LooseTolerance_ReturnsClassC()
        {
            // 500mm nominal, ±6mm band → fits Class C
            var cls = ClassifyLinearTolerance13920(500, 12.0);
            Assert.AreEqual(Iso13920Linear.C, cls);
        }

        [TestMethod]
        public void ClassifyGeometric_TightFlatness_ReturnsClassE()
        {
            // 500mm, 1mm flatness → Class E (1.5mm)
            var cls = ClassifyGeometricTolerance13920(500, 1);
            Assert.AreEqual(Iso13920Geometric.E, cls);
        }

        // =====================================================================
        // Unit conversion
        // =====================================================================

        [TestMethod]
        public void InchesToMm_OneInch()
        {
            Assert.AreEqual(25.4, InchesToMm(1.0), 0.001);
        }

        [TestMethod]
        public void MmToInches_254mm()
        {
            Assert.AreEqual(10.0, MmToInches(254), 0.001);
        }
    }

    [TestClass]
    public class FabricationToleranceClassifierTests
    {
        private readonly FabricationToleranceClassifier _classifier = new FabricationToleranceClassifier();

        // =====================================================================
        // ISO Standard Detection
        // =====================================================================

        [TestMethod]
        public void Classify_DetectsIso13920BF()
        {
            string text = "GENERAL TOLERANCES PER ISO 13920-BF";
            var result = _classifier.Classify(text, null, null);
            Assert.IsTrue(result.Iso13920Detected);
            Assert.AreEqual(Iso13920Linear.B, result.DetectedLinearClass);
            Assert.AreEqual(Iso13920Geometric.F, result.DetectedGeometricClass);
            Assert.IsFalse(result.LinearTighterThanShop);
            Assert.IsFalse(result.GeometricTighterThanShop);
        }

        [TestMethod]
        public void Classify_DetectsIso13920AE_FlagsTighter()
        {
            string text = "TOLERANCE: ISO 13920-AE";
            var result = _classifier.Classify(text, null, null);
            Assert.IsTrue(result.Iso13920Detected);
            Assert.AreEqual(Iso13920Linear.A, result.DetectedLinearClass);
            Assert.AreEqual(Iso13920Geometric.E, result.DetectedGeometricClass);
            Assert.IsTrue(result.LinearTighterThanShop);
            Assert.IsTrue(result.GeometricTighterThanShop);
        }

        [TestMethod]
        public void Classify_DetectsIso13920_LinearOnly()
        {
            string text = "PER ISO 13920-B";
            var result = _classifier.Classify(text, null, null);
            Assert.IsTrue(result.Iso13920Detected);
            Assert.AreEqual(Iso13920Linear.B, result.DetectedLinearClass);
            Assert.IsFalse(result.DetectedGeometricClass.HasValue);
        }

        [TestMethod]
        public void Classify_DetectsIso2768()
        {
            string text = "GENERAL TOLERANCES ISO 2768-m";
            var result = _classifier.Classify(text, null, null);
            Assert.IsTrue(result.Iso2768Detected);
            Assert.AreEqual(Iso2768Linear.Medium, result.DetectedIso2768Class);
        }

        [TestMethod]
        public void Classify_DetectsIso2768Fine()
        {
            string text = "ISO 2768-f";
            var result = _classifier.Classify(text, null, null);
            Assert.IsTrue(result.Iso2768Detected);
            Assert.AreEqual(Iso2768Linear.Fine, result.DetectedIso2768Class);
        }

        // =====================================================================
        // Dimension Classification for Fabrication
        // =====================================================================

        [TestMethod]
        public void Classify_MachiningTolerance_FlagsCritical()
        {
            // ±0.005" = machining territory for a fab shop
            var tolResult = new ToleranceAnalysisResult();
            tolResult.SpecificTolerances.Add(new DimensionTolerance
            {
                Nominal = 2.0,
                Plus = 0.005,
                Minus = 0.005,
                TotalBand = 0.010,
                RawText = "2.000 ±0.005"
            });

            var result = _classifier.Classify("", tolResult, null);
            Assert.IsTrue(result.RequiresMachining);
            Assert.AreEqual(FabricationTier.PrecisionMachining, result.OverallTier);
            Assert.IsTrue(result.CostFlags.Any(f => f.Impact == CostImpact.Critical));
        }

        [TestMethod]
        public void Classify_TightMachiningTolerance_FlagsMachining()
        {
            // ±0.010" = still machining for fab shop
            var tolResult = new ToleranceAnalysisResult();
            tolResult.SpecificTolerances.Add(new DimensionTolerance
            {
                Nominal = 5.0,
                Plus = 0.010,
                Minus = 0.010,
                TotalBand = 0.020,
                RawText = "5.000 ±0.010"
            });

            var result = _classifier.Classify("", tolResult, null);
            Assert.IsTrue(result.RequiresMachining);
            Assert.AreEqual(FabricationTier.Machining, result.OverallTier);
        }

        [TestMethod]
        public void Classify_FabStandardTolerance_NoCostFlags()
        {
            // ±0.060" (~±1.5mm) on a 2" part → ISO 13920 Class B territory
            var tolResult = new ToleranceAnalysisResult();
            tolResult.SpecificTolerances.Add(new DimensionTolerance
            {
                Nominal = 2.0,
                Plus = 0.060,
                Minus = 0.060,
                TotalBand = 0.120,
                RawText = "2.000 ±0.060"
            });

            var result = _classifier.Classify("", tolResult, null);
            Assert.IsFalse(result.RequiresMachining);
            Assert.AreEqual(FabricationTier.ShopStandard, result.OverallTier);
        }

        // =====================================================================
        // GD&T Classification for Fabrication
        // =====================================================================

        [TestMethod]
        public void Classify_TightFlatness_RequiresMachining()
        {
            // 0.005" flatness = 0.127mm → precision machining for fab
            var gdt = new List<GdtCallout>
            {
                new GdtCallout
                {
                    GdtFeatureType = GdtType.Flatness,
                    ToleranceValue = 0.005,
                    RawText = "FLATNESS .005"
                }
            };

            var result = _classifier.Classify("", null, gdt);
            Assert.IsTrue(result.RequiresMachining);
            Assert.IsTrue(result.GdtClassifications.Any(g =>
                g.FabTier >= FabricationTier.PrecisionMachining));
        }

        [TestMethod]
        public void Classify_LooseFlatness_ShopStandard()
        {
            // 0.125" flatness = 3.175mm → within Class F for typical fab part
            var gdt = new List<GdtCallout>
            {
                new GdtCallout
                {
                    GdtFeatureType = GdtType.Flatness,
                    ToleranceValue = 0.125,
                    RawText = "FLATNESS .125"
                }
            };

            var result = _classifier.Classify("", null, gdt);
            Assert.IsFalse(result.RequiresMachining);
            Assert.AreEqual(FabricationTier.ShopStandard, result.OverallTier);
        }

        // =====================================================================
        // Press Brake Bend Stackup
        // =====================================================================

        [TestMethod]
        public void Classify_ManyBendsWithRefs_HighStackupRisk()
        {
            string text = @"
                BEND 1: 90° UP
                BEND 2: 90° DOWN
                BEND 3: 45° UP
                BEND 4: 90° DOWN
                BEND 5: 90° UP
                BEND 6: 135° DOWN
                2.500 FROM BEND LINE
                1.250 BETWEEN BEND TO BEND
            ";

            var result = _classifier.Classify(text, null, null);
            Assert.AreEqual(BendStackupRisk.High, result.BendStackupRisk);
            Assert.IsTrue(result.BendCount >= 6);
            Assert.IsTrue(result.BendRefDimCount >= 2);
        }

        [TestMethod]
        public void Classify_FewBends_LowRisk()
        {
            string text = "BEND 1: 90° UP\nBEND 2: 90° DOWN\nBEND 3: 90° UP";
            var result = _classifier.Classify(text, null, null);
            Assert.AreEqual(BendStackupRisk.Low, result.BendStackupRisk);
        }

        [TestMethod]
        public void Classify_NoBends_NoRisk()
        {
            string text = "FLAT PLATE - NO BENDS";
            var result = _classifier.Classify(text, null, null);
            Assert.AreEqual(BendStackupRisk.None, result.BendStackupRisk);
        }

        // =====================================================================
        // Routing Hints
        // =====================================================================

        [TestMethod]
        public void ToRoutingHints_MachiningRequired_AddsHint()
        {
            var tolResult = new ToleranceAnalysisResult();
            tolResult.SpecificTolerances.Add(new DimensionTolerance
            {
                Nominal = 1.0,
                Plus = 0.005,
                Minus = 0.005,
                TotalBand = 0.010,
                RawText = "1.000 ±0.005"
            });

            var fabResult = _classifier.Classify("", tolResult, null);
            var hints = _classifier.ToRoutingHints(fabResult);
            Assert.IsTrue(hints.Any(h => h.NoteText.Contains("MACHINING REQUIRED")));
        }

        [TestMethod]
        public void ToRoutingHints_HighBendStackup_AddsHint()
        {
            string text = @"
                BEND 1 BEND 2 BEND 3 BEND 4 BEND 5 BEND 6
                DIM FROM BEND LINE
                DIM BETWEEN BEND TO OTHER
            ";

            var fabResult = _classifier.Classify(text, null, null);
            var hints = _classifier.ToRoutingHints(fabResult);
            Assert.IsTrue(hints.Any(h => h.NoteText.Contains("PRESS BRAKE STACKUP")));
        }

        // =====================================================================
        // ISO 13920 BF = shop standard → no cost flags
        // =====================================================================

        [TestMethod]
        public void Classify_ShopStandardBF_NoCostIncrease()
        {
            string text = "WELDING TOLERANCES PER ISO 13920-BF";
            var result = _classifier.Classify(text, null, null);

            Assert.IsTrue(result.Iso13920Detected);
            Assert.IsFalse(result.LinearTighterThanShop);
            Assert.IsFalse(result.GeometricTighterThanShop);

            // Should have a cost flag but with Standard impact (informational)
            var isoFlag = result.CostFlags.FirstOrDefault(f => f.Source == "ISO 13920 comparison");
            Assert.IsNotNull(isoFlag);
            Assert.AreEqual(CostImpact.None, isoFlag.Impact);
        }

        [TestMethod]
        public void Classify_TighterThanShopAE_CostsMore()
        {
            string text = "TOLERANCES: ISO 13920-AE";
            var result = _classifier.Classify(text, null, null);

            Assert.IsTrue(result.LinearTighterThanShop);
            Assert.IsTrue(result.GeometricTighterThanShop);

            var isoFlag = result.CostFlags.FirstOrDefault(f => f.Source == "ISO 13920 comparison");
            Assert.IsNotNull(isoFlag);
            Assert.AreEqual(CostImpact.High, isoFlag.Impact);
        }

        // =====================================================================
        // Summary
        // =====================================================================

        [TestMethod]
        public void Summary_ContainsIsoInfo()
        {
            string text = "ISO 13920-BF";
            var result = _classifier.Classify(text, null, null);
            Assert.IsTrue(result.Summary.Contains("ISO 13920"));
        }

        [TestMethod]
        public void Summary_NullText_NoIso()
        {
            var result = _classifier.Classify(null, null, null);
            Assert.IsTrue(result.Summary.Contains("No ISO standard specified"));
        }

        // =====================================================================
        // Custom shop standard
        // =====================================================================

        [TestMethod]
        public void CustomShopStandard_CF_TreatsB_AsTighter()
        {
            // Shop that normally works to CF (coarser than BF)
            var classifier = new FabricationToleranceClassifier(
                Iso13920Linear.C, Iso13920Geometric.F);

            string text = "ISO 13920-BF";
            var result = classifier.Classify(text, null, null);

            // B is tighter than C → should flag
            Assert.IsTrue(result.LinearTighterThanShop);
        }

        // =====================================================================
        // Realistic drawing scenario
        // =====================================================================

        [TestMethod]
        public void Classify_RealisticFabDrawing()
        {
            string text = @"
                NORTHERN MANUFACTURING
                WELDMENT ASSEMBLY
                MATERIAL: A36 STEEL
                TOLERANCES: ISO 13920-BF
                NOTES:
                1. ALL WELDS PER AWS D1.1
                2. REMOVE ALL BURRS AND SHARP EDGES
                3. BEND RADIUS = 0.060 MIN
                BEND 1: 90° UP
                BEND 2: 90° DOWN
                BEND 3: 45° UP
                1.500 FROM BEND LINE
            ";

            var tolResult = new ToleranceAnalysisResult();
            // One loose tolerance (fab-friendly)
            tolResult.SpecificTolerances.Add(new DimensionTolerance
            {
                Nominal = 12.0, Plus = 0.060, Minus = 0.060,
                TotalBand = 0.120, RawText = "12.000 ±0.060"
            });

            var result = _classifier.Classify(text, tolResult, null);

            Assert.IsTrue(result.Iso13920Detected);
            Assert.IsFalse(result.LinearTighterThanShop);
            Assert.IsFalse(result.RequiresMachining);
            Assert.AreEqual(FabricationTier.ShopStandard, result.OverallTier);
            Assert.AreEqual(BendStackupRisk.Low, result.BendStackupRisk);
        }

        [TestMethod]
        public void Classify_DrawingWithMachiningDims()
        {
            string text = @"
                BRACKET ASSEMBLY
                ISO 13920-AE
                NOTES:
                1. MACHINE MOUNTING HOLES AFTER WELDING
                BEND 1: 90° UP
                BEND 2: 90° DOWN
            ";

            var tolResult = new ToleranceAnalysisResult();
            // One tight machining tolerance
            tolResult.SpecificTolerances.Add(new DimensionTolerance
            {
                Nominal = 0.500, Plus = 0.002, Minus = 0.002,
                TotalBand = 0.004, RawText = "0.500 ±0.002"
            });
            // One standard fab tolerance
            tolResult.SpecificTolerances.Add(new DimensionTolerance
            {
                Nominal = 6.0, Plus = 0.060, Minus = 0.060,
                TotalBand = 0.120, RawText = "6.000 ±0.060"
            });

            var result = _classifier.Classify(text, tolResult, null);

            Assert.IsTrue(result.Iso13920Detected);
            Assert.IsTrue(result.LinearTighterThanShop); // AE > BF
            Assert.IsTrue(result.RequiresMachining); // ±0.002"
            Assert.AreEqual(FabricationTier.PrecisionMachining, result.OverallTier);
            Assert.IsTrue(result.CostFlags.Count >= 2); // ISO flag + machining flag
        }
    }
}
