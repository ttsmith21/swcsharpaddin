using Xunit;
using NM.Core.Pdf;

namespace NM.Core.Tests.Pdf
{
    public class TitleBlockParserTests
    {
        private readonly TitleBlockParser _parser = new TitleBlockParser();

        [Theory]
        [InlineData("PART NO: 12345-A", "12345-A")]
        [InlineData("PART NUMBER: NM-BRACKET-001", "NM-BRACKET-001")]
        [InlineData("DWG NO. 55-1234-REV", "55-1234-REV")]
        [InlineData("P/N: ABC123", "ABC123")]
        [InlineData("DRAWING NUMBER: D-9876", "D-9876")]
        [InlineData("ITEM NO: 42", "42")]
        public void ParsesPartNumber(string input, string expected)
        {
            var result = _parser.Parse(input);
            Assert.Equal(expected, result.PartNumber);
        }

        [Theory]
        [InlineData("MATERIAL: 304 STAINLESS STEEL", "304 STAINLESS STEEL")]
        [InlineData("MAT'L: A36", "A36")]
        [InlineData("MATL: 1018 CRS", "1018 CRS")]
        [InlineData("MATERIAL: ASTM A36 FINISH: PAINT", "ASTM A36")]
        public void ParsesMaterial(string input, string expected)
        {
            var result = _parser.Parse(input);
            Assert.Equal(expected, result.Material);
        }

        [Theory]
        [InlineData("1018 CRS", "1018 CRS")]
        [InlineData("ASTM A-36 STEEL", "ASTM A-36")]
        [InlineData("304 STAINLESS", "304 STAINLESS")]
        [InlineData("6061 ALUMINUM", "6061 ALUMINUM")]
        public void ParsesMaterialCallouts(string input, string expected)
        {
            var result = _parser.Parse(input);
            Assert.NotNull(result.Material);
            Assert.Contains(expected, result.Material, System.StringComparison.OrdinalIgnoreCase);
        }

        [Theory]
        [InlineData("REV: C", "C")]
        [InlineData("REVISION: B", "B")]
        [InlineData("REV A", "A")]
        [InlineData("REV. 3", "3")]
        public void ParsesRevision(string input, string expected)
        {
            var result = _parser.Parse(input);
            Assert.Equal(expected, result.Revision);
        }

        [Theory]
        [InlineData("DESCRIPTION: MOUNTING BRACKET", "MOUNTING BRACKET")]
        [InlineData("TITLE: SUPPORT PLATE", "SUPPORT PLATE")]
        [InlineData("NAME: GUSSET ASSEMBLY", "GUSSET ASSEMBLY")]
        public void ParsesDescription(string input, string expected)
        {
            var result = _parser.Parse(input);
            Assert.Equal(expected, result.Description);
        }

        [Fact]
        public void ParsesDrawnBy()
        {
            var result = _parser.Parse("DRAWN BY: J. SMITH");
            Assert.Equal("J. SMITH", result.DrawnBy);
        }

        [Fact]
        public void ParsesScale()
        {
            var result = _parser.Parse("SCALE: 1:2");
            Assert.Equal("1:2", result.Scale);
        }

        [Fact]
        public void ParsesDate()
        {
            var result = _parser.Parse("DATE: 01/15/2025");
            Assert.NotNull(result.Date);
            Assert.Equal(2025, result.Date.Value.Year);
        }

        [Fact]
        public void ParsesSheet()
        {
            var result = _parser.Parse("SHEET 1 OF 3");
            Assert.Equal("1 OF 3", result.Sheet);
        }

        [Fact]
        public void ParsesCompleteTitleBlock()
        {
            string titleBlock = @"
                PART NO: NM-1234-A
                DESCRIPTION: MOUNTING BRACKET
                MATERIAL: 304 STAINLESS STEEL
                FINISH: #4 BRUSHED
                REV: C
                DRAWN BY: J. SMITH
                DATE: 01/15/2025
                SCALE: 1:1
                SHEET 1 OF 2
            ";

            var result = _parser.Parse(titleBlock);

            Assert.Equal("NM-1234-A", result.PartNumber);
            Assert.Equal("MOUNTING BRACKET", result.Description);
            Assert.Contains("304 STAINLESS STEEL", result.Material);
            Assert.Equal("C", result.Revision);
            Assert.Equal("J. SMITH", result.DrawnBy);
            Assert.Equal("1:1", result.Scale);
            Assert.Equal("1 OF 2", result.Sheet);
            Assert.True(result.OverallConfidence > 0.5);
        }

        [Fact]
        public void EmptyText_ReturnsEmptyInfo()
        {
            var result = _parser.Parse("");
            Assert.Null(result.PartNumber);
            Assert.Null(result.Material);
            Assert.Null(result.Revision);
            Assert.Equal(0, result.OverallConfidence);
        }

        [Fact]
        public void NullText_ReturnsEmptyInfo()
        {
            var result = _parser.Parse(null);
            Assert.Null(result.PartNumber);
        }
    }
}
