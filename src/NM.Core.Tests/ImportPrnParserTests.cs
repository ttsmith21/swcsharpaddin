using System.Linq;
using NM.Core.Export;
using Xunit;

namespace NM.Core.Tests
{
    public class ImportPrnParserTests
    {
        [Fact]
        public void ParseLines_ImSection_ExtractsPartNumber()
        {
            var lines = new[]
            {
                "DECL(IM) ADD IM-KEY IM-DRAWING IM-DESCR IM-REV IM-TYPE IM-CLASS IM-COMMODITY IM-STD-LOT",
                "END",
                "\"B1_NativeBracket_14ga_CS\" \"\" \"304L BENT\" \"\" 1 9 \"F\" 1"
            };

            var records = ImportPrnParser.ParseLines(lines);

            Assert.True(records.ContainsKey("B1_NativeBracket_14ga_CS"));
            var rec = records["B1_NativeBracket_14ga_CS"];
            Assert.Equal("304L BENT", rec.Description);
            Assert.Equal(1, rec.ImType);
            Assert.Equal(9, rec.ImClass);
        }

        [Fact]
        public void ParseLines_EmptyQuotedStrings_PreservedAsTokens()
        {
            // Verify that "" fields don't shift alignment
            var lines = new[]
            {
                "DECL(IM) ADD IM-KEY IM-DRAWING IM-DESCR IM-REV IM-TYPE IM-CLASS IM-COMMODITY IM-STD-LOT",
                "END",
                "\"TestPart\" \"\" \"\" \"\" 1 9 \"F\" 2"
            };

            var records = ImportPrnParser.ParseLines(lines);

            Assert.True(records.ContainsKey("TestPart"));
            var rec = records["TestPart"];
            Assert.Equal("", rec.Drawing);
            Assert.Equal("", rec.Description);
            Assert.Equal("", rec.Revision);
            Assert.Equal(1, rec.ImType);
            Assert.Equal(9, rec.ImClass);
            Assert.Equal(2, rec.StdLot);
        }

        [Fact]
        public void ParseLines_EscapedQuotes_HandledCorrectly()
        {
            var lines = new[]
            {
                "DECL(IM) ADD IM-KEY IM-DRAWING IM-DESCR IM-REV IM-TYPE IM-CLASS IM-COMMODITY IM-STD-LOT",
                "END",
                "\"D10_Pipe_Reducer\" \"\" \"3\\\"X1-1/2\\\" 304SS STRAIGHT REDUCER\" \"\" 1 9 \"F\" 1"
            };

            var records = ImportPrnParser.ParseLines(lines);

            Assert.True(records.ContainsKey("D10_Pipe_Reducer"));
            var rec = records["D10_Pipe_Reducer"];
            Assert.Equal("3\"X1-1/2\" 304SS STRAIGHT REDUCER", rec.Description);
        }

        [Fact]
        public void ParseLines_BomSection_ExtractsQuantity()
        {
            var lines = new[]
            {
                "DECL(PS) ADD PS-PARENT-KEY PS-SUBORD-KEY PS-REV PS-PIECE-NO PS-QTY-P",
                "END",
                "\"GOLD_STANDARD_ASM_CLEAN\" \"B1_NativeBracket_14ga_CS\" \"COMMON SET\" \"07\" 1"
            };

            var records = ImportPrnParser.ParseLines(lines);

            Assert.True(records.ContainsKey("B1_NativeBracket_14ga_CS"));
            var rec = records["B1_NativeBracket_14ga_CS"];
            Assert.Equal("GOLD_STANDARD_ASM_CLEAN", rec.ParentPartNumber);
            Assert.Equal("07", rec.PieceNumber);
            Assert.Equal(1.0, rec.BomQuantity);
        }

        [Fact]
        public void ParseLines_MaterialSection_ExtractsOptiMaterial()
        {
            var lines = new[]
            {
                "DECL(PS) ADD PS-PARENT-KEY PS-SUBORD-KEY PS-REV PS-PIECE-NO PS-QTY-P PS-DIM-1 PS-ISSUE-SW PS-BFLOCATION-SW PS-BFQTY-SW PS-BFZEROQTY-SW PS-OP-NUM",
                "END",
                "\"B1_NativeBracket_14ga_CS\" \"S.304L14GA\" \"COMMON SET\" \"01\" 12.4348 0 2 1 1 1 20"
            };

            var records = ImportPrnParser.ParseLines(lines);

            Assert.True(records.ContainsKey("B1_NativeBracket_14ga_CS"));
            var rec = records["B1_NativeBracket_14ga_CS"];
            Assert.Equal("S.304L14GA", rec.OptiMaterial);
            Assert.Equal(12.4348, rec.RawWeight, 4);
            Assert.Equal(0.0, rec.F300Length);
        }

        [Fact]
        public void ParseLines_MaterialWithEscapedQuotes_ParsesStockId()
        {
            var lines = new[]
            {
                "DECL(PS) ADD PS-PARENT-KEY PS-SUBORD-KEY PS-REV PS-PIECE-NO PS-QTY-P PS-DIM-1 PS-ISSUE-SW PS-BFLOCATION-SW PS-BFQTY-SW PS-BFZEROQTY-SW PS-OP-NUM",
                "END",
                "\"C1_RoundTube_2OD_SCH40\" \"P.304L12\\\"SCH40S\" \"COMMON SET\" \"01\" 1. 6 2 1 1 1 20"
            };

            var records = ImportPrnParser.ParseLines(lines);

            Assert.True(records.ContainsKey("C1_RoundTube_2OD_SCH40"));
            var rec = records["C1_RoundTube_2OD_SCH40"];
            Assert.Equal("P.304L12\"SCH40S", rec.OptiMaterial);
            Assert.Equal(1.0, rec.RawWeight);
            Assert.Equal(6.0, rec.F300Length);
        }

        [Fact]
        public void ParseLines_RoutingSection_ExtractsWorkCenters()
        {
            var lines = new[]
            {
                "DECL(RT) ADD RT-ITEM-KEY RT-WORKCENTER-KEY RT-OP-NUM RT-SETUP RT-RUN-STD RT-REV RT-MULT-SEQ",
                "END",
                "\"B1_NativeBracket_14ga_CS\" \"O120\" 10 0 0 \"COMMON SET\" 0",
                "\"B1_NativeBracket_14ga_CS\" \"N120\" 20 .01 .008 \"COMMON SET\" 0",
                "\"B1_NativeBracket_14ga_CS\" \"N140\" 40 .18 .0125 \"COMMON SET\" 0"
            };

            var records = ImportPrnParser.ParseLines(lines);

            Assert.True(records.ContainsKey("B1_NativeBracket_14ga_CS"));
            var rec = records["B1_NativeBracket_14ga_CS"];
            Assert.Equal(3, rec.Routing.Count);

            var n120 = rec.Routing.First(r => r.WorkCenter == "N120");
            Assert.Equal(20, n120.OpNumber);
            Assert.Equal(0.01, n120.SetupHours);
            Assert.Equal(0.008, n120.RunHours);

            var n140 = rec.Routing.First(r => r.WorkCenter == "N140");
            Assert.Equal(40, n140.OpNumber);
            Assert.Equal(0.18, n140.SetupHours);
            Assert.Equal(0.0125, n140.RunHours);
        }

        [Fact]
        public void ParseLines_RoutingNotes_ExtractsNotes()
        {
            var lines = new[]
            {
                "DECL(RN) ADD RN-ITEM-KEY RN-OP-NUM RN-LINE-NO RN-REV RN-DESCR",
                "END",
                "\"B1_NativeBracket_14ga_CS\" 20 1 \"COMMON SET\" \"AUTHOR: tschoen\"",
                "\"B1_NativeBracket_14ga_CS\" 20 2 \"COMMON SET\" \"DATE: 2/6/2026\"",
                "\"B1_NativeBracket_14ga_CS\" 40 1 \"COMMON SET\" \"BRAKE TO CAD\""
            };

            var records = ImportPrnParser.ParseLines(lines);

            Assert.True(records.ContainsKey("B1_NativeBracket_14ga_CS"));
            var rec = records["B1_NativeBracket_14ga_CS"];
            Assert.Equal(2, rec.RoutingNotes.Count); // ops 20 and 40
            Assert.Equal(2, rec.RoutingNotes[20].Count);
            Assert.Equal("AUTHOR: tschoen", rec.RoutingNotes[20][0]);
            Assert.Single(rec.RoutingNotes[40]);
            Assert.Equal("BRAKE TO CAD", rec.RoutingNotes[40][0]);
        }

        [Fact]
        public void ParseLines_MultipleDecl_CombinesData()
        {
            // Simulate IM + BOM PS + Material PS + RT all contributing to the same record
            var lines = new[]
            {
                "DECL(IM) ADD IM-KEY IM-DRAWING IM-DESCR IM-REV IM-TYPE IM-CLASS IM-COMMODITY IM-STD-LOT",
                "END",
                "\"B1_NativeBracket_14ga_CS\" \"\" \"304L BENT\" \"\" 1 9 \"F\" 1",
                "",
                "DECL(PS) ADD PS-PARENT-KEY PS-SUBORD-KEY PS-REV PS-PIECE-NO PS-QTY-P",
                "END",
                "\"GOLD_STANDARD_ASM_CLEAN\" \"B1_NativeBracket_14ga_CS\" \"COMMON SET\" \"07\" 1",
                "",
                "DECL(PS) ADD PS-PARENT-KEY PS-SUBORD-KEY PS-REV PS-PIECE-NO PS-QTY-P PS-DIM-1 PS-ISSUE-SW PS-BFLOCATION-SW PS-BFQTY-SW PS-BFZEROQTY-SW PS-OP-NUM",
                "END",
                "\"B1_NativeBracket_14ga_CS\" \"S.304L14GA\" \"COMMON SET\" \"01\" 12.4348 0 2 1 1 1 20",
                "",
                "DECL(RT) ADD RT-ITEM-KEY RT-WORKCENTER-KEY RT-OP-NUM RT-SETUP RT-RUN-STD RT-REV RT-MULT-SEQ",
                "END",
                "\"B1_NativeBracket_14ga_CS\" \"N120\" 20 .01 .008 \"COMMON SET\" 0"
            };

            var records = ImportPrnParser.ParseLines(lines);

            var rec = records["B1_NativeBracket_14ga_CS"];
            Assert.Equal("304L BENT", rec.Description);
            Assert.Equal("S.304L14GA", rec.OptiMaterial);
            Assert.Equal(12.4348, rec.RawWeight, 4);
            Assert.Equal(1.0, rec.BomQuantity);
            Assert.Equal("GOLD_STANDARD_ASM_CLEAN", rec.ParentPartNumber);
            Assert.Single(rec.Routing);
            Assert.Equal("N120", rec.Routing[0].WorkCenter);
        }

        [Fact]
        public void ParseLines_EmptyInput_ReturnsEmptyDictionary()
        {
            var records = ImportPrnParser.ParseLines(new string[0]);
            Assert.Empty(records);
        }

        [Fact]
        public void ParseLines_BlankLinesOnly_ReturnsEmptyDictionary()
        {
            var lines = new[] { "", "  ", "\t", "" };
            var records = ImportPrnParser.ParseLines(lines);
            Assert.Empty(records);
        }

        [Fact]
        public void ParseLines_PsDifferentiation_BomVsMaterial()
        {
            // Two PS sections: one with PS-DIM-1 (material), one without (BOM)
            var lines = new[]
            {
                "DECL(PS) ADD PS-PARENT-KEY PS-SUBORD-KEY PS-REV PS-PIECE-NO PS-QTY-P",
                "END",
                "\"ASSY\" \"PART1\" \"REV\" \"01\" 3",
                "",
                "DECL(PS) ADD PS-PARENT-KEY PS-SUBORD-KEY PS-REV PS-PIECE-NO PS-QTY-P PS-DIM-1 PS-ISSUE-SW PS-BFLOCATION-SW PS-BFQTY-SW PS-BFZEROQTY-SW PS-OP-NUM",
                "END",
                "\"PART1\" \"S.STOCK123\" \"REV\" \"01\" 50.5 12.3 2 1 1 1 20"
            };

            var records = ImportPrnParser.ParseLines(lines);

            var rec = records["PART1"];
            Assert.Equal(3.0, rec.BomQuantity);         // From BOM PS section
            Assert.Equal("ASSY", rec.ParentPartNumber);  // From BOM PS section
            Assert.Equal("S.STOCK123", rec.OptiMaterial); // From Material PS section
            Assert.Equal(50.5, rec.RawWeight);            // From Material PS section
            Assert.Equal(12.3, rec.F300Length);            // From Material PS section
        }
    }
}
