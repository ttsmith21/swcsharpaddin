using Xunit;
using NM.Core.DataModel;
using NM.Core.Processing;

namespace NM.Core.Tests
{
    public class DescriptionGeneratorTests
    {
        [Fact]
        public void SheetMetal_WithBends_ReturnsBent()
        {
            var pd = new PartData
            {
                Classification = PartType.SheetMetal,
                Material = "304L"
            };
            pd.Sheet.IsSheetMetal = true;
            pd.Sheet.BendCount = 3;

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("304L BENT", result);
        }

        [Fact]
        public void SheetMetal_WithF325RollForming_ReturnsRoll()
        {
            var pd = new PartData
            {
                Classification = PartType.SheetMetal,
                Material = "304L"
            };
            pd.Sheet.IsSheetMetal = true;
            pd.Sheet.BendCount = 2;
            pd.Cost.F325_Price = 15.50;

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("304L ROLL", result);
        }

        [Fact]
        public void SheetMetal_WithLargeBendRadius_ReturnsRoll()
        {
            var pd = new PartData
            {
                Classification = PartType.SheetMetal,
                Material = "304L"
            };
            pd.Sheet.IsSheetMetal = true;
            pd.Sheet.BendCount = 1;
            pd.Extra["MaxBendRadiusIn"] = "12.5000";

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("304L ROLL", result);
        }

        [Fact]
        public void SheetMetal_Flat_NoBends_ReturnsPlate()
        {
            var pd = new PartData
            {
                Classification = PartType.SheetMetal,
                Material = "304L"
            };
            pd.Sheet.IsSheetMetal = true;
            pd.Sheet.BendCount = 0;

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("304L PLATE", result);
        }

        [Fact]
        public void Tube_Round_ReturnsPipe()
        {
            var pd = new PartData
            {
                Classification = PartType.Tube,
                Material = "304L"
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "Round";

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("304L PIPE", result);
        }

        [Fact]
        public void Tube_Square_ReturnsSqTube()
        {
            var pd = new PartData
            {
                Classification = PartType.Tube,
                Material = "304L"
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "Square";

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("304L SQ TUBE", result);
        }

        [Fact]
        public void Tube_Rectangular_ReturnsRectTube()
        {
            var pd = new PartData
            {
                Classification = PartType.Tube,
                Material = "304L"
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "Rectangle";

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("304L RECT TUBE", result);
        }

        [Fact]
        public void Tube_Angle_ReturnsAngle()
        {
            var pd = new PartData
            {
                Classification = PartType.Tube,
                Material = "304L"
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "Angle";

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("304L ANGLE", result);
        }

        [Fact]
        public void NullPartData_ReturnsNull()
        {
            string result = DescriptionGenerator.Generate(null);

            Assert.Null(result);
        }

        [Fact]
        public void EmptyMaterial_ReturnsNull()
        {
            var pd = new PartData
            {
                Classification = PartType.SheetMetal,
                Material = ""
            };

            string result = DescriptionGenerator.Generate(pd);

            Assert.Null(result);
        }

        [Fact]
        public void WhitespaceMaterial_ReturnsNull()
        {
            var pd = new PartData
            {
                Classification = PartType.SheetMetal,
                Material = "   "
            };

            string result = DescriptionGenerator.Generate(pd);

            Assert.Null(result);
        }

        [Fact]
        public void GenericPart_ReturnsMaterialOnly()
        {
            var pd = new PartData
            {
                Classification = PartType.Generic,
                Material = "A36"
            };

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("A36", result);
        }

        [Fact]
        public void Tube_Channel_ReturnsChannel()
        {
            var pd = new PartData
            {
                Classification = PartType.Tube,
                Material = "A36"
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "Channel";

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("A36 CHANNEL", result);
        }

        [Fact]
        public void MaterialIsUppercased()
        {
            var pd = new PartData
            {
                Classification = PartType.Generic,
                Material = "a36"
            };

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("A36", result);
        }

        [Fact]
        public void SwMaterialName_AISI304_MapsTo304L()
        {
            var pd = new PartData
            {
                Classification = PartType.SheetMetal,
                Material = "AISI 304"
            };
            pd.Sheet.IsSheetMetal = true;
            pd.Sheet.BendCount = 2;

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("304L BENT", result);
        }

        [Fact]
        public void SwMaterialName_AISI316_MapsTo316L()
        {
            var pd = new PartData
            {
                Classification = PartType.Tube,
                Material = "AISI 316"
            };
            pd.Tube.IsTube = true;
            pd.Tube.TubeShape = "Round";

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("316L PIPE", result);
        }

        [Fact]
        public void SwMaterialName_PlainCarbonSteel_MapsToCS()
        {
            var pd = new PartData
            {
                Classification = PartType.SheetMetal,
                Material = "Plain Carbon Steel"
            };
            pd.Sheet.IsSheetMetal = true;
            pd.Sheet.BendCount = 0;

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("CS PLATE", result);
        }

        [Fact]
        public void ExtraMaterialCode_TakesPriorityOverSwName()
        {
            var pd = new PartData
            {
                Classification = PartType.SheetMetal,
                Material = "AISI 304"
            };
            pd.Sheet.IsSheetMetal = true;
            pd.Sheet.BendCount = 1;
            pd.Extra["MaterialCode"] = "304L";

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("304L BENT", result);
        }

        [Fact]
        public void UnknownSwMaterial_PassesThroughUppercased()
        {
            var pd = new PartData
            {
                Classification = PartType.Generic,
                Material = "Inconel 625"
            };

            string result = DescriptionGenerator.Generate(pd);

            Assert.Equal("INCONEL 625", result);
        }
    }
}
