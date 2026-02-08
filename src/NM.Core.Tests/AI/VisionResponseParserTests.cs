using Xunit;
using NM.Core.AI;
using NM.Core.AI.Models;

namespace NM.Core.Tests.AI
{
    public class VisionResponseParserTests
    {
        private readonly VisionResponseParser _parser = new VisionResponseParser();

        [Fact]
        public void ParsesTitleBlockResponse()
        {
            string json = @"{
                ""part_number"": ""NM-1234-A"",
                ""description"": ""MOUNTING BRACKET"",
                ""revision"": ""C"",
                ""material"": ""304 STAINLESS STEEL"",
                ""finish"": ""#4 BRUSHED"",
                ""drawn_by"": ""J. SMITH"",
                ""date"": ""01/15/2025"",
                ""scale"": ""1:1"",
                ""sheet"": ""1 OF 2"",
                ""tolerance_general"": ""+/- .010""
            }";

            var result = _parser.ParseTitleBlockResponse(json);

            Assert.True(result.Success);
            Assert.Equal("NM-1234-A", result.PartNumber.Value);
            Assert.Equal("MOUNTING BRACKET", result.Description.Value);
            Assert.Equal("C", result.Revision.Value);
            Assert.Equal("304 STAINLESS STEEL", result.Material.Value);
            Assert.Equal("#4 BRUSHED", result.Finish.Value);
            Assert.Equal("J. SMITH", result.DrawnBy.Value);
            Assert.Equal("1:1", result.Scale.Value);
            Assert.Equal("1 OF 2", result.Sheet.Value);
        }

        [Fact]
        public void ParsesFullPageResponse()
        {
            string json = @"{
                ""title_block"": {
                    ""part_number"": ""55-1234"",
                    ""description"": ""GUSSET PLATE"",
                    ""revision"": ""B"",
                    ""material"": ""A36"",
                    ""finish"": """"
                },
                ""dimensions"": {
                    ""overall_length_inches"": 12.5,
                    ""overall_width_inches"": 6.25,
                    ""thickness_inches"": 0.25
                },
                ""manufacturing_notes"": [
                    {
                        ""text"": ""BREAK ALL EDGES"",
                        ""category"": ""deburr"",
                        ""routing_impact"": ""add_operation""
                    },
                    {
                        ""text"": ""GALVANIZE PER ASTM A123"",
                        ""category"": ""finish"",
                        ""routing_impact"": ""add_operation""
                    }
                ],
                ""gdt_callouts"": [
                    {
                        ""type"": ""flatness"",
                        ""tolerance"": ""0.010"",
                        ""datum_references"": [""A""],
                        ""feature_description"": ""Bottom surface""
                    }
                ],
                ""holes"": {
                    ""tapped_holes"": [""1/4-20""],
                    ""through_holes"": [""0.500 DIA THRU""]
                },
                ""bend_info"": {
                    ""bend_radius"": ""0.125"",
                    ""bend_count"": ""3""
                },
                ""special_requirements"": [""DFAR COMPLIANT""]
            }";

            var result = _parser.ParseFullPageResponse(json);

            Assert.True(result.Success);
            Assert.Equal("55-1234", result.PartNumber.Value);
            Assert.Equal("GUSSET PLATE", result.Description.Value);
            Assert.Equal("B", result.Revision.Value);
            Assert.Equal("A36", result.Material.Value);

            // Dimensions
            Assert.NotNull(result.OverallLength);
            Assert.Equal("12.5000", result.OverallLength.Value);
            Assert.NotNull(result.Thickness);
            Assert.Equal("0.2500", result.Thickness.Value);

            // Notes
            Assert.Equal(2, result.ManufacturingNotes.Count);
            Assert.Equal("BREAK ALL EDGES", result.ManufacturingNotes[0].Text);
            Assert.Equal("deburr", result.ManufacturingNotes[0].Category);

            // GD&T
            Assert.Single(result.GdtCallouts);
            Assert.Equal("flatness", result.GdtCallouts[0].Type);
            Assert.Equal("0.010", result.GdtCallouts[0].Tolerance);
            Assert.Contains("A", result.GdtCallouts[0].DatumReferences);

            // Holes
            Assert.Single(result.TappedHoles);
            Assert.Equal("1/4-20", result.TappedHoles[0]);
            Assert.Single(result.ThroughHoles);

            // Bends
            Assert.Equal("0.125", result.BendRadius.Value);
            Assert.Equal("3", result.BendCount.Value);

            // Special requirements
            Assert.Single(result.SpecialRequirements);
            Assert.Equal("DFAR COMPLIANT", result.SpecialRequirements[0]);
        }

        [Fact]
        public void StripsMarkdownFencing()
        {
            string json = "```json\n{\"part_number\": \"12345\"}\n```";
            var result = _parser.ParseTitleBlockResponse(json);
            Assert.True(result.Success);
            Assert.Equal("12345", result.PartNumber.Value);
        }

        [Fact]
        public void HandlesEmptyJson()
        {
            var result = _parser.ParseTitleBlockResponse("{}");
            Assert.True(result.Success);
            Assert.False(result.PartNumber?.HasValue ?? false);
        }

        [Fact]
        public void HandlesNullInput()
        {
            var result = _parser.ParseTitleBlockResponse(null);
            Assert.False(result.Success);
            Assert.Contains("Failed to parse", result.ErrorMessage);
        }

        [Fact]
        public void HandlesMalformedJson()
        {
            var result = _parser.ParseTitleBlockResponse("not json at all");
            Assert.False(result.Success);
        }

        [Fact]
        public void HandlesPartialResponse()
        {
            string json = @"{""part_number"": ""12345"", ""material"": ""304 SS""}";
            var result = _parser.ParseTitleBlockResponse(json);
            Assert.True(result.Success);
            Assert.Equal("12345", result.PartNumber.Value);
            Assert.Equal("304 SS", result.Material.Value);
            Assert.False(result.Revision?.HasValue ?? false);
        }

        [Fact]
        public void HandlesNullFields()
        {
            string json = @"{""part_number"": ""12345"", ""material"": null, ""revision"": null}";
            var result = _parser.ParseTitleBlockResponse(json);
            Assert.True(result.Success);
            Assert.Equal("12345", result.PartNumber.Value);
            Assert.False(result.Material?.HasValue ?? false);
        }

        [Fact]
        public void ExtractsJsonFromSurroundingText()
        {
            string response = "Here is the analysis:\n{\"part_number\": \"ABC-123\"}\nEnd of analysis.";
            var result = _parser.ParseTitleBlockResponse(response);
            Assert.True(result.Success);
            Assert.Equal("ABC-123", result.PartNumber.Value);
        }

        [Fact]
        public void OverallConfidence_CalculatesCorrectly()
        {
            string json = @"{
                ""part_number"": ""12345"",
                ""material"": ""A36"",
                ""revision"": ""A""
            }";
            var result = _parser.ParseTitleBlockResponse(json);
            Assert.True(result.OverallConfidence > 0);
        }

        [Fact]
        public void OverallConfidence_ZeroForEmptyResult()
        {
            var result = _parser.ParseTitleBlockResponse("{}");
            Assert.Equal(0, result.OverallConfidence);
        }
    }
}
