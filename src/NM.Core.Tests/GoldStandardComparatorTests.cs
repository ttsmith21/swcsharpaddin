using System.Collections.Generic;
using NM.Core.Testing;
using Xunit;

namespace NM.Core.Tests
{
    public class GoldStandardComparatorTests
    {
        [Fact]
        public void FieldComparison_DefaultStatus_IsFail()
        {
            var fc = new FieldComparison();
            Assert.Equal(MatchStatus.Fail, fc.Status);
        }

        [Fact]
        public void PartComparisonResult_CalculateCounts_SumsCorrectly()
        {
            var result = new PartComparisonResult { FileName = "test.sldprt" };
            result.Fields.Add(new FieldComparison { Status = MatchStatus.Match });
            result.Fields.Add(new FieldComparison { Status = MatchStatus.Match });
            result.Fields.Add(new FieldComparison { Status = MatchStatus.TolerancePass });
            result.Fields.Add(new FieldComparison { Status = MatchStatus.NotImplemented });
            result.Fields.Add(new FieldComparison { Status = MatchStatus.Fail });

            result.CalculateCounts();

            Assert.Equal(2, result.MatchCount);
            Assert.Equal(1, result.TolerancePassCount);
            Assert.Equal(1, result.NotImplementedCount);
            Assert.Equal(1, result.FailCount);
            Assert.Equal(MatchStatus.Fail, result.OverallStatus);
        }

        [Fact]
        public void PartComparisonResult_AllMatch_OverallStatusIsMatch()
        {
            var result = new PartComparisonResult { FileName = "test.sldprt" };
            result.Fields.Add(new FieldComparison { Status = MatchStatus.Match });
            result.Fields.Add(new FieldComparison { Status = MatchStatus.Match });

            result.CalculateCounts();

            Assert.Equal(MatchStatus.Match, result.OverallStatus);
        }

        [Fact]
        public void PartComparisonResult_NotImplementedWorstStatus_OverallNotImplemented()
        {
            var result = new PartComparisonResult { FileName = "test.sldprt" };
            result.Fields.Add(new FieldComparison { Status = MatchStatus.Match });
            result.Fields.Add(new FieldComparison { Status = MatchStatus.NotImplemented });

            result.CalculateCounts();

            Assert.Equal(MatchStatus.NotImplemented, result.OverallStatus);
        }

        [Fact]
        public void FullComparisonReport_CalculateTotals_AggregatesAcrossParts()
        {
            var report = new FullComparisonReport();

            var part1 = new PartComparisonResult { FileName = "part1.sldprt" };
            part1.Fields.Add(new FieldComparison { Status = MatchStatus.Match });
            part1.Fields.Add(new FieldComparison { Status = MatchStatus.Fail });
            part1.CalculateCounts();

            var part2 = new PartComparisonResult { FileName = "part2.sldprt" };
            part2.Fields.Add(new FieldComparison { Status = MatchStatus.Match });
            part2.Fields.Add(new FieldComparison { Status = MatchStatus.NotImplemented });
            part2.CalculateCounts();

            report.Parts.Add(part1);
            report.Parts.Add(part2);
            report.CalculateTotals();

            Assert.Equal(4, report.TotalFields);
            Assert.Equal(2, report.TotalMatch);
            Assert.Equal(1, report.TotalFail);
            Assert.Equal(1, report.TotalNotImplemented);
        }

        [Fact]
        public void FullComparisonReport_GenerateSummary_ProducesNonEmptyText()
        {
            var report = new FullComparisonReport();

            var part = new PartComparisonResult { FileName = "test.sldprt" };
            part.Fields.Add(new FieldComparison { Status = MatchStatus.Match, FieldName = "Classification" });
            part.CalculateCounts();
            report.Parts.Add(part);
            report.CalculateTotals();

            var summary = report.GenerateSummary();

            Assert.Contains("GOLD STANDARD COMPARISON SUMMARY", summary);
            Assert.Contains("Match", summary);
        }

        [Fact]
        public void FullComparisonReport_GenerateDetailedReport_IncludesPartNames()
        {
            var report = new FullComparisonReport();

            var part = new PartComparisonResult { FileName = "B1_NativeBracket.sldprt" };
            part.Fields.Add(new FieldComparison
            {
                FieldName = "Classification",
                ExpectedValue = "SheetMetal",
                ActualValue = "SheetMetal",
                Status = MatchStatus.Match
            });
            part.Fields.Add(new FieldComparison
            {
                FieldName = "OptiMaterial",
                ExpectedValue = "S.304L14GA",
                ActualValue = "",
                Status = MatchStatus.NotImplemented,
                Note = "Not yet wired"
            });
            part.CalculateCounts();
            report.Parts.Add(part);
            report.CalculateTotals();

            var detailed = report.GenerateDetailedReport();

            Assert.Contains("B1_NativeBracket", detailed);
            Assert.Contains("OptiMaterial", detailed);
            Assert.Contains("NOT IMPLEMENTED", detailed);
        }

        [Fact]
        public void CompareFromData_BasicMatch_ReportsCorrectly()
        {
            var manifest = new Dictionary<string, object>
            {
                { "version", "2.0" },
                { "defaultTolerances", new Dictionary<string, object>
                    {
                        { "thickness", 0.001 },
                        { "cost", 0.05 }
                    }
                },
                { "files", new Dictionary<string, object>
                    {
                        { "test.sldprt", new Dictionary<string, object>
                            {
                                { "shouldPass", true },
                                { "expectedClassification", "SheetMetal" },
                                { "expectedThickness_in", 0.073 },
                                { "vbaBaseline", new Dictionary<string, object>() },
                                { "csharpExpected", new Dictionary<string, object>() },
                                { "knownDeviations", new Dictionary<string, object>() }
                            }
                        }
                    }
                }
            };

            var results = new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    { "FileName", "test.sldprt" },
                    { "Classification", "SheetMetal" },
                    { "Thickness_in", 0.073 }
                }
            };

            var report = GoldStandardComparator.CompareFromData(manifest, results);

            Assert.Single(report.Parts);
            var part = report.Parts[0];
            Assert.Equal("test.sldprt", part.FileName);

            // Classification should match
            var classField = part.Fields.Find(f => f.FieldName == "Classification");
            Assert.NotNull(classField);
            Assert.Equal(MatchStatus.Match, classField.Status);
        }

        [Fact]
        public void CompareFromData_KnownDeviation_MarkedAsNotImplemented()
        {
            var manifest = new Dictionary<string, object>
            {
                { "version", "2.0" },
                { "defaultTolerances", new Dictionary<string, object>() },
                { "files", new Dictionary<string, object>
                    {
                        { "test.sldprt", new Dictionary<string, object>
                            {
                                { "shouldPass", true },
                                { "vbaBaseline", new Dictionary<string, object>
                                    {
                                        { "optiMaterial", "S.304L14GA" }
                                    }
                                },
                                { "csharpExpected", new Dictionary<string, object>() },
                                { "knownDeviations", new Dictionary<string, object>
                                    {
                                        { "optiMaterial", new Dictionary<string, object>
                                            {
                                                { "reason", "Not yet wired" },
                                                { "status", "NOT_IMPLEMENTED" }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            };

            var results = new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    { "FileName", "test.sldprt" },
                    { "OptiMaterial", "" }
                }
            };

            var report = GoldStandardComparator.CompareFromData(manifest, results);
            var part = report.Parts[0];
            var optiField = part.Fields.Find(f => f.FieldName == "OptiMaterial");
            Assert.NotNull(optiField);
            Assert.Equal(MatchStatus.NotImplemented, optiField.Status);
        }

        [Fact]
        public void CompareFromData_CsharpOverridesTakePrecedence()
        {
            var manifest = new Dictionary<string, object>
            {
                { "version", "2.0" },
                { "defaultTolerances", new Dictionary<string, object>() },
                { "files", new Dictionary<string, object>
                    {
                        { "test.sldprt", new Dictionary<string, object>
                            {
                                { "shouldPass", true },
                                { "vbaBaseline", new Dictionary<string, object>
                                    {
                                        { "description", "304L BENT" }
                                    }
                                },
                                { "csharpExpected", new Dictionary<string, object>
                                    {
                                        { "description", "304L BENT BRACKET" }
                                    }
                                },
                                { "knownDeviations", new Dictionary<string, object>() }
                            }
                        }
                    }
                }
            };

            var results = new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    { "FileName", "test.sldprt" },
                    { "Description", "304L BENT BRACKET" }
                }
            };

            var report = GoldStandardComparator.CompareFromData(manifest, results);
            var part = report.Parts[0];
            var descField = part.Fields.Find(f => f.FieldName == "Description");
            Assert.NotNull(descField);
            Assert.Equal(MatchStatus.Match, descField.Status);
        }
    }
}
