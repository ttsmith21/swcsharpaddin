using NM.Core.AI.Models;

namespace NM.Core.AI
{
    /// <summary>
    /// Structured prompts for AI drawing analysis.
    /// Separated from the service so they can be tested and tuned independently.
    /// </summary>
    public static class VisionPrompts
    {
        /// <summary>
        /// Prompt for title block extraction only (Tier 2).
        /// Returns a small, focused JSON structure.
        /// </summary>
        public static string GetTitleBlockPrompt()
        {
            return @"Analyze this engineering drawing title block and extract information as JSON.
Return ONLY valid JSON, no other text, no markdown fencing.

{
  ""part_number"": """",
  ""description"": """",
  ""revision"": """",
  ""material"": """",
  ""finish"": """",
  ""drawn_by"": """",
  ""date"": """",
  ""scale"": """",
  ""sheet"": """",
  ""tolerance_general"": """"
}

Rules:
- Use empty string """" for fields you cannot determine.
- For material, include the full specification (e.g., ""ASTM A36"", ""304 STAINLESS STEEL"", ""1018 CRS"").
- For revision, include only the letter or number (e.g., ""C"", ""3"").
- For part_number, exclude revision suffixes.
- For date, use MM/DD/YYYY format.
- For sheet, use ""X OF Y"" format if multiple sheets.";
        }

        /// <summary>
        /// Prompt for full drawing page analysis (Tier 3).
        /// Returns a comprehensive JSON structure with notes, GD&amp;T, and geometry.
        /// </summary>
        public static string GetFullPagePrompt()
        {
            return @"You are a manufacturing engineer analyzing an engineering drawing for a metal fabrication shop.
Extract ALL manufacturing-relevant information as JSON.
Return ONLY valid JSON, no other text, no markdown fencing.

{
  ""title_block"": {
    ""part_number"": """",
    ""description"": """",
    ""revision"": """",
    ""material"": """",
    ""finish"": """",
    ""drawn_by"": """",
    ""date"": """",
    ""scale"": """",
    ""sheet"": """",
    ""tolerance_general"": """"
  },
  ""dimensions"": {
    ""overall_length_inches"": null,
    ""overall_width_inches"": null,
    ""thickness_inches"": null
  },
  ""manufacturing_notes"": [
    {
      ""text"": ""exact note text as written on drawing"",
      ""category"": ""deburr|finish|heat_treat|weld|machine|inspect|hardware|process_constraint|general"",
      ""routing_impact"": ""add_operation|modify_operation|informational|constraint""
    }
  ],
  ""gdt_callouts"": [
    {
      ""type"": ""flatness|parallelism|perpendicularity|position|profile|runout|concentricity"",
      ""tolerance"": """",
      ""datum_references"": [],
      ""feature_description"": """"
    }
  ],
  ""holes"": {
    ""tapped_holes"": [],
    ""through_holes"": []
  },
  ""bend_info"": {
    ""bend_radius"": """",
    ""bend_count"": """"
  },
  ""special_requirements"": []
}

Rules:
- Only include data you can clearly read from the drawing.
- For dimensions, convert to inches if drawing uses metric.
- For notes, preserve the exact wording from the drawing.
- Categorize each note by its manufacturing impact.
- For tapped_holes, use format like ""1/4-20"", ""M6x1.0"".
- For through_holes, use format like ""0.250 DIA THRU"".
- For special_requirements, include items like ITAR, DFAR, MIL-SPEC callouts, PPAP, etc.";
        }

        /// <summary>
        /// Prompt for full page analysis with known context from the 3D model.
        /// Allows the AI to cross-validate and flag discrepancies.
        /// </summary>
        public static string GetFullPagePromptWithContext(VisionContext context)
        {
            string basePrompt = GetFullPagePrompt();

            if (context == null || !context.HasContext)
                return basePrompt;

            string contextBlock = "\n\nADDITIONAL CONTEXT FROM 3D MODEL (use for cross-validation):\n";

            if (!string.IsNullOrEmpty(context.KnownPartNumber))
                contextBlock += $"- 3D model part number: {context.KnownPartNumber}\n";
            if (!string.IsNullOrEmpty(context.KnownMaterial))
                contextBlock += $"- 3D model material: {context.KnownMaterial}\n";
            if (context.KnownThickness_in.HasValue)
                contextBlock += $"- 3D model thickness: {context.KnownThickness_in.Value:F4} inches\n";
            if (!string.IsNullOrEmpty(context.KnownClassification))
                contextBlock += $"- Part classification: {context.KnownClassification}\n";

            contextBlock += @"
If the drawing data DISAGREES with the 3D model data above, add a ""discrepancies"" array to the JSON:
""discrepancies"": [
  {
    ""field"": ""material"",
    ""drawing_value"": ""what drawing says"",
    ""model_value"": ""what 3D model says"",
    ""recommendation"": ""which value is likely correct and why""
  }
]";

            return basePrompt + contextBlock;
        }

        /// <summary>
        /// System prompt establishing the AI's role.
        /// </summary>
        public static string GetSystemPrompt()
        {
            return @"You are an expert manufacturing engineer specializing in sheet metal fabrication,
tube processing, and CNC machining. You analyze engineering drawings to extract
manufacturing-relevant data for ERP routing and quoting systems.

You are precise, conservative, and never guess. If you cannot clearly read a field,
leave it as empty string. You always return valid JSON with no extra text.";
        }
    }
}
