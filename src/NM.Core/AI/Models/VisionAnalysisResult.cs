using System;
using System.Collections.Generic;

namespace NM.Core.AI.Models
{
    /// <summary>
    /// Structured result from AI vision analysis of a drawing page.
    /// </summary>
    public sealed class VisionAnalysisResult
    {
        // Title block fields
        public FieldResult PartNumber { get; set; }
        public FieldResult Description { get; set; }
        public FieldResult Revision { get; set; }
        public FieldResult Material { get; set; }
        public FieldResult Finish { get; set; }
        public FieldResult DrawnBy { get; set; }
        public FieldResult Date { get; set; }
        public FieldResult Scale { get; set; }
        public FieldResult Sheet { get; set; }
        public FieldResult ToleranceGeneral { get; set; }

        // Geometry hints
        public FieldResult OverallLength { get; set; }
        public FieldResult OverallWidth { get; set; }
        public FieldResult Thickness { get; set; }

        // Manufacturing notes (full page analysis only)
        public List<NoteResult> ManufacturingNotes { get; } = new List<NoteResult>();

        // GD&T callouts (full page analysis only)
        public List<GdtResult> GdtCallouts { get; } = new List<GdtResult>();

        // Hole/feature data (full page analysis only)
        public List<string> TappedHoles { get; } = new List<string>();
        public List<string> ThroughHoles { get; } = new List<string>();

        // Bend info (sheet metal drawings)
        public FieldResult BendRadius { get; set; }
        public FieldResult BendCount { get; set; }

        // Special requirements
        public List<string> SpecialRequirements { get; } = new List<string>();

        // Metadata
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string RawJson { get; set; }
        public decimal CostUsd { get; set; }
        public int InputTokens { get; set; }
        public int OutputTokens { get; set; }
        public TimeSpan Duration { get; set; }

        /// <summary>Overall confidence based on how many fields were extracted.</summary>
        public double OverallConfidence
        {
            get
            {
                int populated = 0;
                int total = 0;
                double sumConfidence = 0;

                CheckField(PartNumber, ref populated, ref total, ref sumConfidence);
                CheckField(Description, ref populated, ref total, ref sumConfidence);
                CheckField(Revision, ref populated, ref total, ref sumConfidence);
                CheckField(Material, ref populated, ref total, ref sumConfidence);
                CheckField(Finish, ref populated, ref total, ref sumConfidence);

                return total > 0 ? sumConfidence / total : 0;
            }
        }

        private static void CheckField(FieldResult field, ref int populated, ref int total, ref double sum)
        {
            total++;
            if (field != null && !string.IsNullOrEmpty(field.Value))
            {
                populated++;
                sum += field.Confidence;
            }
        }
    }

    /// <summary>
    /// A single extracted field with confidence score.
    /// </summary>
    public sealed class FieldResult
    {
        public string Value { get; set; }
        public double Confidence { get; set; }

        public FieldResult() { }

        public FieldResult(string value, double confidence = 0.9)
        {
            Value = value;
            Confidence = confidence;
        }

        /// <summary>True if the field has a non-empty value.</summary>
        public bool HasValue => !string.IsNullOrEmpty(Value);

        public override string ToString() => Value ?? "";
    }

    /// <summary>
    /// A manufacturing note extracted by AI.
    /// </summary>
    public sealed class NoteResult
    {
        public string Text { get; set; }
        public string Category { get; set; }
        public string RoutingImpact { get; set; }
        public double Confidence { get; set; }
    }

    /// <summary>
    /// A GD&amp;T callout extracted by AI.
    /// </summary>
    public sealed class GdtResult
    {
        public string Type { get; set; }
        public string Tolerance { get; set; }
        public List<string> DatumReferences { get; set; } = new List<string>();
        public string FeatureDescription { get; set; }
        public double Confidence { get; set; }
    }
}
