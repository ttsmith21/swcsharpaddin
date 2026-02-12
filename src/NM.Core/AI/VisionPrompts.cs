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

        // ========================================================================
        // Multi-pass focused prompts (Phase 2B)
        // Each pass has a focused schema reducing hallucination opportunity.
        // ========================================================================

        /// <summary>
        /// Pass 1: Title block extraction from cropped title block region.
        /// Focused on 10 fields only, high accuracy on small image.
        /// </summary>
        public static string GetPass1TitleBlockPrompt()
        {
            return @"Extract ONLY the title block information from this engineering drawing image.
Return ONLY valid JSON, no other text.

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

IMPORTANT RULES:
- ONLY extract data visible in the title block area.
- Use empty string for fields you cannot clearly read.
- material: Full spec (e.g., ""ASTM A36"", ""304 STAINLESS STEEL"", ""6061-T6 ALUMINUM"").
- revision: Letter or number only (e.g., ""C"", ""3""), NOT revision history.
- part_number: The primary identifier, exclude revision suffixes.
- tolerance_general: The ""UNLESS OTHERWISE SPECIFIED"" tolerance block text.
- Do NOT infer or guess any values.";
        }

        /// <summary>
        /// Pass 2: Manufacturing notes scan across the full drawing page.
        /// Focused on finding all notes and categorizing them.
        /// </summary>
        public static string GetPass2NotesPrompt()
        {
            return @"Scan this engineering drawing for ALL manufacturing notes, callouts, and process instructions.
Return ONLY valid JSON, no other text.

{
  ""manufacturing_notes"": [
    {
      ""text"": ""exact note text as written on drawing"",
      ""category"": ""deburr|finish|heat_treat|weld|machine|inspect|hardware|process_constraint|material|general"",
      ""routing_impact"": ""add_operation|modify_operation|informational|constraint"",
      ""location"": ""notes_section|view_annotation|general_note|title_block""
    }
  ],
  ""special_requirements"": [],
  ""holes"": {
    ""tapped_holes"": [],
    ""through_holes"": []
  },
  ""bend_info"": {
    ""bend_radius"": """",
    ""bend_count"": """"
  }
}

RULES:
- Check ALL areas: NOTES section, view annotations, flag notes, general notes.
- Preserve EXACT wording from the drawing.
- Do NOT include dimension callouts as notes.
- Do NOT include title block information as notes.
- For tapped_holes: ""1/4-20"", ""M6x1.0"" format.
- For through_holes: ""0.250 DIA THRU"" format.
- special_requirements: ITAR, DFAR, MIL-SPEC, PPAP, first article, etc.
- If no notes are found, return empty arrays.";
        }

        /// <summary>
        /// Pass 3: Tolerance scan — identifies tolerances tighter than the general tolerance.
        /// Requires the general tolerance to be provided as context.
        /// </summary>
        public static string GetPass3TolerancePrompt(string generalToleranceText = null)
        {
            string contextLine = "";
            if (!string.IsNullOrEmpty(generalToleranceText))
            {
                contextLine = $@"
The drawing's general tolerance block states: ""{generalToleranceText}""
Only report tolerances that are TIGHTER than these general defaults.
";
            }
            else
            {
                contextLine = @"
If you can read the general tolerance block (UNLESS OTHERWISE SPECIFIED section),
extract it first, then identify any tolerances tighter than those defaults.
";
            }

            return @"Scan this engineering drawing for GD&T feature control frames and tolerances
that are TIGHTER than the general (default) tolerance.
Return ONLY valid JSON, no other text.
" + contextLine + @"
{
  ""general_tolerance"": {
    ""raw_text"": """",
    ""two_place_decimal"": null,
    ""three_place_decimal"": null,
    ""four_place_decimal"": null,
    ""angles_degrees"": null,
    ""fractional"": null
  },
  ""tight_tolerances"": [
    {
      ""dimension_value"": """",
      ""tolerance"": """",
      ""tolerance_type"": ""bilateral|unilateral|limit|basic"",
      ""feature_description"": """",
      ""is_tighter_than_general"": true
    }
  ],
  ""gdt_callouts"": [
    {
      ""type"": ""flatness|parallelism|perpendicularity|position|profile_of_surface|profile_of_line|circular_runout|total_runout|concentricity|symmetry"",
      ""tolerance"": """",
      ""datum_references"": [],
      ""feature_description"": """",
      ""modifier"": ""none|MMC|LMC|RFS""
    }
  ]
}

RULES:
- ONLY report specific callouts that deviate from the general tolerance.
- For bilateral: ""+/-0.005"" or ""+-0.005"".
- For limit dimensions: ""1.000/1.002"".
- For GD&T, identify the geometric characteristic symbol accurately.
- datum_references: list of datum letters (e.g., [""A"", ""B"", ""C""]).
- If no tight tolerances or GD&T callouts exist, return empty arrays.
- Do NOT hallucinate tolerances that are not explicitly shown.";
        }

        // ========================================================================
        // Rename Wizard prompts
        // ========================================================================

        /// <summary>
        /// Extracts a BOM table from an assembly drawing PDF page.
        /// Returns a JSON array of line items with part number, description, material, quantity.
        /// </summary>
        public static string GetBomTablePrompt()
        {
            return @"Extract the Bill of Materials (BOM) table from this engineering drawing.
Return ONLY valid JSON, no other text, no markdown fencing.

{
  ""bom_items"": [
    {
      ""item_number"": 1,
      ""part_number"": """",
      ""description"": """",
      ""material"": """",
      ""quantity"": 1
    }
  ]
}

RULES:
- Extract ALL rows from the BOM table exactly as shown.
- item_number: The item/find number column (integer).
- part_number: The part number or drawing number column. Preserve exact text.
- description: The description or name column. Preserve exact text.
- material: Material specification if present in the BOM, otherwise empty string.
- quantity: Numeric quantity (integer). Default to 1 if unclear.
- Do NOT skip any rows.
- Do NOT infer or guess values not visible in the BOM table.
- If no BOM table is present on this page, return an empty array.
- Ignore revision columns, weight columns, and other non-essential fields.";
        }

        /// <summary>
        /// Asks AI to match BOM rows to STEP-imported component names.
        /// Provides both lists as context and expects a mapping back.
        /// </summary>
        public static string GetBomMatchingPrompt(string bomJson, string componentsJson)
        {
            return @"You are matching BOM (Bill of Materials) line items to STEP-imported component filenames.

BOM items extracted from the drawing:
" + bomJson + @"

Component filenames from the assembly:
" + componentsJson + @"

Return ONLY valid JSON, no other text, no markdown fencing.

{
  ""matches"": [
    {
      ""component_index"": 0,
      ""bom_item_number"": 1,
      ""predicted_name"": ""the part_number or description from the BOM"",
      ""confidence"": 0.95,
      ""reason"": ""brief explanation of why this match was chosen""
    }
  ]
}

RULES:
- Match each component to the most likely BOM item based on:
  1. Name similarity (STEP filenames often contain partial part numbers).
  2. Quantity correlation (component instance count vs BOM quantity).
  3. Description keywords appearing in the filename.
- confidence: 0.0 to 1.0. Use 0.95+ only for very clear matches.
- If a component has NO reasonable BOM match, omit it from the matches array.
- predicted_name: Use the BOM part_number. If part_number is empty, use description.
- Do NOT force matches. It is better to omit a component than match it incorrectly.
- Multiple components CAN map to the same BOM item (e.g., multiple instances).";
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
