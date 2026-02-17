using System;
using System.Collections.Generic;

namespace NM.Core.Manufacturing
{
    /// <summary>
    /// One row from the WPS lookup table (CSV/spreadsheet).
    /// Each entry represents a qualified Welding Procedure Specification.
    /// </summary>
    public sealed class WpsEntry
    {
        /// <summary>WPS number, e.g., "WPS-001" or "NM-GMAW-SS-001".</summary>
        public string WpsNumber { get; set; }

        /// <summary>Welding process code, e.g., "GMAW", "GTAW", "FCAW", "SMAW".</summary>
        public string Process { get; set; }

        /// <summary>Base metal 1 category: "CS", "SS", "AL".</summary>
        public string BaseMetal1 { get; set; }

        /// <summary>Base metal 2 category: "CS", "SS", "AL". Same as BaseMetal1 for similar joints.</summary>
        public string BaseMetal2 { get; set; }

        /// <summary>Minimum qualified thickness in inches (inclusive).</summary>
        public double ThicknessMinIn { get; set; }

        /// <summary>Maximum qualified thickness in inches (inclusive).</summary>
        public double ThicknessMaxIn { get; set; }

        /// <summary>Joint type: "Groove", "Fillet", "Both".</summary>
        public string JointType { get; set; }

        /// <summary>Applicable welding code: "D1.1", "D1.6", "ASME IX", etc.</summary>
        public string Code { get; set; }

        /// <summary>Filler metal specification, e.g., "ER70S-6", "ER308L", "ER4043".</summary>
        public string FillerMetal { get; set; }

        /// <summary>Shielding gas, e.g., "75/25 Ar/CO2", "100% Ar".</summary>
        public string ShieldingGas { get; set; }

        /// <summary>Optional notes or restrictions.</summary>
        public string Notes { get; set; }
    }

    /// <summary>
    /// Input describing a weld joint for WPS lookup.
    /// Built from pipeline data (materials, thicknesses) and drawing symbols.
    /// </summary>
    public sealed class WpsJointInput
    {
        /// <summary>Material category of first member: "CS", "SS", "AL".</summary>
        public string BaseMetal1 { get; set; }

        /// <summary>Material category of second member: "CS", "SS", "AL". Same as BaseMetal1 for similar joints.</summary>
        public string BaseMetal2 { get; set; }

        /// <summary>Governing thickness in inches (thinner of the two members for groove welds).</summary>
        public double ThicknessIn { get; set; }

        /// <summary>Joint type from drawing symbol: "Groove", "Fillet", or empty for auto-match.</summary>
        public string JointType { get; set; }

        /// <summary>Required welding code: "D1.1", "D1.6", "ASME IX", or empty for any.</summary>
        public string RequiredCode { get; set; }
    }

    /// <summary>
    /// Result of WPS resolution for one joint.
    /// </summary>
    public sealed class WpsMatchResult
    {
        /// <summary>Matched WPS entries (best match first). Empty if no match found.</summary>
        public List<WpsEntry> MatchedEntries { get; } = new List<WpsEntry>();

        /// <summary>True when at least one WPS matched.</summary>
        public bool HasMatch => MatchedEntries.Count > 0;

        /// <summary>Best-match WPS number, or empty if no match.</summary>
        public string WpsNumber => HasMatch ? MatchedEntries[0].WpsNumber : string.Empty;

        /// <summary>Flags indicating conditions that need human review.</summary>
        public List<WpsReviewFlag> ReviewFlags { get; } = new List<WpsReviewFlag>();

        /// <summary>True when any condition requires human review before proceeding.</summary>
        public bool NeedsReview => ReviewFlags.Count > 0;

        /// <summary>Human-readable summary of the match or review reasons.</summary>
        public string Summary { get; set; }
    }

    /// <summary>
    /// Reasons a weld joint is flagged for human review.
    /// </summary>
    public enum WpsReviewReason
    {
        /// <summary>Groove weld on material thicker than 1/2" (12.7mm).</summary>
        ThickGrooveWeld,

        /// <summary>Dissimilar base metals (e.g., CS to SS).</summary>
        DissimilarMetals,

        /// <summary>No qualified WPS found in the lookup table.</summary>
        NoMatchingWps,

        /// <summary>Multiple WPS entries match — engineer must select.</summary>
        AmbiguousMatch,

        /// <summary>Aluminum welding — always verify procedure qualification.</summary>
        AluminumWeld
    }

    /// <summary>
    /// One review flag with reason and description.
    /// </summary>
    public sealed class WpsReviewFlag
    {
        public WpsReviewReason Reason { get; set; }
        public string Description { get; set; }

        public WpsReviewFlag() { }

        public WpsReviewFlag(WpsReviewReason reason, string description)
        {
            Reason = reason;
            Description = description;
        }
    }
}
