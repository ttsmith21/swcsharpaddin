using System.Collections.Generic;
using System.Linq;
using NM.Core.Manufacturing;
using Xunit;

namespace NM.Core.Tests
{
    public class WpsResolverTests
    {
        private static WpsLookupTable BuildTestTable()
        {
            return WpsLookupTable.FromEntries(new[]
            {
                new WpsEntry { WpsNumber = "WPS-CS-001", Process = "GMAW", BaseMetal1 = "CS", BaseMetal2 = "CS", ThicknessMinIn = 0.036, ThicknessMaxIn = 0.500, JointType = "Both", Code = "D1.1", FillerMetal = "ER70S-6", ShieldingGas = "75/25" },
                new WpsEntry { WpsNumber = "WPS-CS-002", Process = "GMAW", BaseMetal1 = "CS", BaseMetal2 = "CS", ThicknessMinIn = 0.500, ThicknessMaxIn = 1.500, JointType = "Groove", Code = "D1.1", FillerMetal = "ER70S-6", ShieldingGas = "75/25" },
                new WpsEntry { WpsNumber = "WPS-SS-001", Process = "GMAW", BaseMetal1 = "SS", BaseMetal2 = "SS", ThicknessMinIn = 0.036, ThicknessMaxIn = 0.500, JointType = "Both", Code = "D1.6", FillerMetal = "ER308L", ShieldingGas = "98/2" },
                new WpsEntry { WpsNumber = "WPS-SS-002", Process = "GTAW", BaseMetal1 = "SS", BaseMetal2 = "SS", ThicknessMinIn = 0.036, ThicknessMaxIn = 0.250, JointType = "Both", Code = "D1.6", FillerMetal = "ER308L", ShieldingGas = "100% Ar" },
                new WpsEntry { WpsNumber = "WPS-AL-001", Process = "GMAW", BaseMetal1 = "AL", BaseMetal2 = "AL", ThicknessMinIn = 0.063, ThicknessMaxIn = 0.500, JointType = "Both", Code = "D1.1", FillerMetal = "ER4043", ShieldingGas = "100% Ar" },
                new WpsEntry { WpsNumber = "WPS-CSSS-001", Process = "GMAW", BaseMetal1 = "CS", BaseMetal2 = "SS", ThicknessMinIn = 0.036, ThicknessMaxIn = 0.500, JointType = "Fillet", Code = "D1.1", FillerMetal = "ER309L", ShieldingGas = "75/25" },
            });
        }

        // ---- Basic matching ----

        [Fact]
        public void CarbonSteel_Fillet_FindsMatch()
        {
            var table = BuildTestTable();
            var result = WpsResolver.ResolveForPart("CS", 0.125, "Fillet", table);

            Assert.True(result.HasMatch);
            Assert.Equal("WPS-CS-001", result.WpsNumber);
            Assert.False(result.NeedsReview);
        }

        [Fact]
        public void StainlessSteel_Groove_ThinGauge_FindsMultipleMatches()
        {
            var table = BuildTestTable();
            var result = WpsResolver.ResolveForPart("SS", 0.125, "Groove", table);

            // Both WPS-SS-001 (Both joint) and WPS-SS-002 (Both joint) match
            Assert.True(result.HasMatch);
            Assert.True(result.MatchedEntries.Count >= 2);
        }

        [Fact]
        public void NoMatchingThickness_ReturnsNoMatch()
        {
            var table = BuildTestTable();
            var result = WpsResolver.ResolveForPart("CS", 3.0, "Groove", table);

            Assert.False(result.HasMatch);
            Assert.True(result.NeedsReview);
            Assert.Contains(result.ReviewFlags, f => f.Reason == WpsReviewReason.NoMatchingWps);
        }

        [Fact]
        public void EmptyJointType_MatchesAll()
        {
            var table = BuildTestTable();
            var result = WpsResolver.ResolveForPart("CS", 0.125, "", table);

            Assert.True(result.HasMatch);
            Assert.Equal("WPS-CS-001", result.WpsNumber);
        }

        // ---- Review flags ----

        [Fact]
        public void ThickGrooveWeld_FlagsForReview()
        {
            var table = BuildTestTable();
            var result = WpsResolver.ResolveForPart("CS", 0.750, "Groove", table);

            Assert.True(result.NeedsReview);
            Assert.Contains(result.ReviewFlags, f => f.Reason == WpsReviewReason.ThickGrooveWeld);
        }

        [Fact]
        public void ThickFilletWeld_DoesNotFlagThickGroove()
        {
            var table = BuildTestTable();
            var result = WpsResolver.ResolveForPart("CS", 0.750, "Fillet", table);

            // Fillet welds don't trigger the thick-groove check
            Assert.DoesNotContain(result.ReviewFlags, f => f.Reason == WpsReviewReason.ThickGrooveWeld);
        }

        [Fact]
        public void DissimilarMetals_FlagsForReview()
        {
            var table = BuildTestTable();
            var result = WpsResolver.ResolveForAssemblyJoint("CS", 0.125, "SS", 0.125, "Fillet", table);

            Assert.True(result.NeedsReview);
            Assert.Contains(result.ReviewFlags, f => f.Reason == WpsReviewReason.DissimilarMetals);
        }

        [Fact]
        public void DissimilarMetals_StillFindsMatchingWps()
        {
            var table = BuildTestTable();
            var result = WpsResolver.ResolveForAssemblyJoint("CS", 0.125, "SS", 0.125, "Fillet", table);

            Assert.True(result.HasMatch);
            Assert.Equal("WPS-CSSS-001", result.WpsNumber);
        }

        [Fact]
        public void AluminumWeld_AlwaysFlaggedForReview()
        {
            var table = BuildTestTable();
            var result = WpsResolver.ResolveForPart("AL", 0.125, "Fillet", table);

            Assert.True(result.NeedsReview);
            Assert.Contains(result.ReviewFlags, f => f.Reason == WpsReviewReason.AluminumWeld);
        }

        [Fact]
        public void NoTableLoaded_FlagsNoMatchingWps()
        {
            var emptyTable = new WpsLookupTable();
            var result = WpsResolver.ResolveForPart("CS", 0.125, "Fillet", emptyTable);

            Assert.True(result.NeedsReview);
            Assert.Contains(result.ReviewFlags, f => f.Reason == WpsReviewReason.NoMatchingWps);
        }

        [Fact]
        public void NullTable_FlagsNoMatchingWps()
        {
            var result = WpsResolver.ResolveForPart("CS", 0.125, "Fillet", null);

            Assert.True(result.NeedsReview);
            Assert.Contains(result.ReviewFlags, f => f.Reason == WpsReviewReason.NoMatchingWps);
        }

        // ---- Assembly joint resolution ----

        [Fact]
        public void AssemblyJoint_UsesGoverningThickness()
        {
            var table = BuildTestTable();
            // Thin member = 0.125, thick member = 0.750
            // Governing thickness should be 0.125 (thinner)
            var result = WpsResolver.ResolveForAssemblyJoint("CS", 0.750, "CS", 0.125, "Fillet", table);

            Assert.True(result.HasMatch);
            Assert.Equal("WPS-CS-001", result.WpsNumber); // 0.125 is within 0.036-0.500 range
        }

        // ---- Bidirectional material matching ----

        [Fact]
        public void MaterialPair_MatchesBidirectionally()
        {
            var table = BuildTestTable();
            // Table has CS+SS, input is SS+CS — should still match
            var result = WpsResolver.ResolveForAssemblyJoint("SS", 0.125, "CS", 0.125, "Fillet", table);

            Assert.True(result.HasMatch);
            Assert.Equal("WPS-CSSS-001", result.WpsNumber);
        }

        // ---- Summary generation ----

        [Fact]
        public void Summary_ContainsMaterialInfo()
        {
            var table = BuildTestTable();
            var result = WpsResolver.ResolveForPart("CS", 0.125, "Fillet", table);

            Assert.False(string.IsNullOrEmpty(result.Summary));
            Assert.Contains("WPS-CS-001", result.Summary);
        }

        [Fact]
        public void ReviewSummary_ContainsReviewReasons()
        {
            var table = BuildTestTable();
            var result = WpsResolver.ResolveForPart("CS", 0.750, "Groove", table);

            Assert.Contains("REVIEW", result.Summary);
        }
    }

    public class WpsLookupTableTests
    {
        [Fact]
        public void LoadFromCsv_HandlesNonexistentFile()
        {
            var table = WpsLookupTable.LoadFromCsv(@"C:\nonexistent\file.csv");
            Assert.False(table.IsLoaded);
            Assert.Empty(table.Entries);
        }

        [Fact]
        public void LoadFromCsv_HandlesNullPath()
        {
            var table = WpsLookupTable.LoadFromCsv(null);
            Assert.False(table.IsLoaded);
        }

        [Fact]
        public void FromEntries_LoadsCorrectly()
        {
            var entries = new[]
            {
                new WpsEntry { WpsNumber = "WPS-001", BaseMetal1 = "CS", BaseMetal2 = "CS", ThicknessMinIn = 0.036, ThicknessMaxIn = 0.5, JointType = "Both" },
                new WpsEntry { WpsNumber = "WPS-002", BaseMetal1 = "SS", BaseMetal2 = "SS", ThicknessMinIn = 0.036, ThicknessMaxIn = 0.5, JointType = "Fillet" },
            };

            var table = WpsLookupTable.FromEntries(entries);
            Assert.True(table.IsLoaded);
            Assert.Equal(2, table.Entries.Count);
        }

        [Fact]
        public void FindMatches_FiltersByThicknessRange()
        {
            var table = WpsLookupTable.FromEntries(new[]
            {
                new WpsEntry { WpsNumber = "THIN", BaseMetal1 = "CS", BaseMetal2 = "CS", ThicknessMinIn = 0.036, ThicknessMaxIn = 0.250, JointType = "Both" },
                new WpsEntry { WpsNumber = "THICK", BaseMetal1 = "CS", BaseMetal2 = "CS", ThicknessMinIn = 0.250, ThicknessMaxIn = 1.500, JointType = "Both" },
            });

            var input = new WpsJointInput { BaseMetal1 = "CS", BaseMetal2 = "CS", ThicknessIn = 0.125 };
            var matches = table.FindMatches(input);

            Assert.Single(matches);
            Assert.Equal("THIN", matches[0].WpsNumber);
        }

        [Fact]
        public void FindMatches_FiltersByJointType()
        {
            var table = WpsLookupTable.FromEntries(new[]
            {
                new WpsEntry { WpsNumber = "FILLET", BaseMetal1 = "CS", BaseMetal2 = "CS", ThicknessMinIn = 0.036, ThicknessMaxIn = 1.0, JointType = "Fillet" },
                new WpsEntry { WpsNumber = "GROOVE", BaseMetal1 = "CS", BaseMetal2 = "CS", ThicknessMinIn = 0.036, ThicknessMaxIn = 1.0, JointType = "Groove" },
            });

            var input = new WpsJointInput { BaseMetal1 = "CS", BaseMetal2 = "CS", ThicknessIn = 0.5, JointType = "Groove" };
            var matches = table.FindMatches(input);

            Assert.Single(matches);
            Assert.Equal("GROOVE", matches[0].WpsNumber);
        }

        [Fact]
        public void FindMatches_BothJointType_MatchesAny()
        {
            var table = WpsLookupTable.FromEntries(new[]
            {
                new WpsEntry { WpsNumber = "BOTH", BaseMetal1 = "CS", BaseMetal2 = "CS", ThicknessMinIn = 0.036, ThicknessMaxIn = 1.0, JointType = "Both" },
            });

            var grooveInput = new WpsJointInput { BaseMetal1 = "CS", BaseMetal2 = "CS", ThicknessIn = 0.5, JointType = "Groove" };
            var filletInput = new WpsJointInput { BaseMetal1 = "CS", BaseMetal2 = "CS", ThicknessIn = 0.5, JointType = "Fillet" };

            Assert.Single(table.FindMatches(grooveInput));
            Assert.Single(table.FindMatches(filletInput));
        }
    }

    public class WpsModelsTests
    {
        [Fact]
        public void WpsMatchResult_EmptyByDefault()
        {
            var result = new WpsMatchResult();
            Assert.False(result.HasMatch);
            Assert.False(result.NeedsReview);
            Assert.Equal(string.Empty, result.WpsNumber);
        }

        [Fact]
        public void WpsMatchResult_HasMatch_WhenEntriesExist()
        {
            var result = new WpsMatchResult();
            result.MatchedEntries.Add(new WpsEntry { WpsNumber = "WPS-001" });
            Assert.True(result.HasMatch);
            Assert.Equal("WPS-001", result.WpsNumber);
        }

        [Fact]
        public void WpsMatchResult_NeedsReview_WhenFlagsExist()
        {
            var result = new WpsMatchResult();
            result.ReviewFlags.Add(new WpsReviewFlag(WpsReviewReason.DissimilarMetals, "test"));
            Assert.True(result.NeedsReview);
        }
    }
}
