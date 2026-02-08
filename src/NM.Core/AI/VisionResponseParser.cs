using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using NM.Core.AI.Models;

namespace NM.Core.AI
{
    /// <summary>
    /// Parses JSON responses from the AI vision service into VisionAnalysisResult.
    /// Handles both title-block-only and full-page response formats.
    /// Resilient to malformed or partial JSON.
    /// </summary>
    public sealed class VisionResponseParser
    {
        /// <summary>
        /// Parses a title-block-only JSON response.
        /// </summary>
        public VisionAnalysisResult ParseTitleBlockResponse(string json)
        {
            var result = new VisionAnalysisResult();

            var obj = SafeParseJson(json);
            if (obj == null)
            {
                result.Success = false;
                result.ErrorMessage = "Failed to parse AI response as JSON";
                result.RawJson = json;
                return result;
            }

            result.Success = true;
            result.RawJson = json;

            result.PartNumber = ExtractField(obj, "part_number");
            result.Description = ExtractField(obj, "description");
            result.Revision = ExtractField(obj, "revision");
            result.Material = ExtractField(obj, "material");
            result.Finish = ExtractField(obj, "finish");
            result.DrawnBy = ExtractField(obj, "drawn_by");
            result.Date = ExtractField(obj, "date");
            result.Scale = ExtractField(obj, "scale");
            result.Sheet = ExtractField(obj, "sheet");
            result.ToleranceGeneral = ExtractField(obj, "tolerance_general");

            return result;
        }

        /// <summary>
        /// Parses a full-page JSON response with notes, GD&amp;T, dimensions, etc.
        /// </summary>
        public VisionAnalysisResult ParseFullPageResponse(string json)
        {
            var result = new VisionAnalysisResult();

            var obj = SafeParseJson(json);
            if (obj == null)
            {
                result.Success = false;
                result.ErrorMessage = "Failed to parse AI response as JSON";
                result.RawJson = json;
                return result;
            }

            result.Success = true;
            result.RawJson = json;

            // Title block (nested object)
            var titleBlock = obj["title_block"] as JObject;
            if (titleBlock != null)
            {
                result.PartNumber = ExtractField(titleBlock, "part_number");
                result.Description = ExtractField(titleBlock, "description");
                result.Revision = ExtractField(titleBlock, "revision");
                result.Material = ExtractField(titleBlock, "material");
                result.Finish = ExtractField(titleBlock, "finish");
                result.DrawnBy = ExtractField(titleBlock, "drawn_by");
                result.Date = ExtractField(titleBlock, "date");
                result.Scale = ExtractField(titleBlock, "scale");
                result.Sheet = ExtractField(titleBlock, "sheet");
                result.ToleranceGeneral = ExtractField(titleBlock, "tolerance_general");
            }

            // Dimensions
            var dimensions = obj["dimensions"] as JObject;
            if (dimensions != null)
            {
                result.OverallLength = ExtractNumericField(dimensions, "overall_length_inches");
                result.OverallWidth = ExtractNumericField(dimensions, "overall_width_inches");
                result.Thickness = ExtractNumericField(dimensions, "thickness_inches");
            }

            // Manufacturing notes
            var notes = obj["manufacturing_notes"] as JArray;
            if (notes != null)
            {
                foreach (var note in notes)
                {
                    var noteObj = note as JObject;
                    if (noteObj == null) continue;

                    result.ManufacturingNotes.Add(new NoteResult
                    {
                        Text = noteObj.Value<string>("text") ?? "",
                        Category = noteObj.Value<string>("category") ?? "general",
                        RoutingImpact = noteObj.Value<string>("routing_impact") ?? "informational",
                        Confidence = 0.85
                    });
                }
            }

            // GD&T callouts
            var gdt = obj["gdt_callouts"] as JArray;
            if (gdt != null)
            {
                foreach (var item in gdt)
                {
                    var gdtObj = item as JObject;
                    if (gdtObj == null) continue;

                    var gdtResult = new GdtResult
                    {
                        Type = gdtObj.Value<string>("type") ?? "",
                        Tolerance = gdtObj.Value<string>("tolerance") ?? "",
                        FeatureDescription = gdtObj.Value<string>("feature_description") ?? "",
                        Confidence = 0.75
                    };

                    var datums = gdtObj["datum_references"] as JArray;
                    if (datums != null)
                    {
                        foreach (var d in datums)
                            gdtResult.DatumReferences.Add(d.ToString());
                    }

                    result.GdtCallouts.Add(gdtResult);
                }
            }

            // Holes
            var holes = obj["holes"] as JObject;
            if (holes != null)
            {
                ExtractStringArray(holes, "tapped_holes", result.TappedHoles);
                ExtractStringArray(holes, "through_holes", result.ThroughHoles);
            }

            // Bend info
            var bendInfo = obj["bend_info"] as JObject;
            if (bendInfo != null)
            {
                result.BendRadius = ExtractField(bendInfo, "bend_radius");
                result.BendCount = ExtractField(bendInfo, "bend_count");
            }

            // Special requirements
            var specials = obj["special_requirements"] as JArray;
            if (specials != null)
            {
                foreach (var s in specials)
                {
                    string val = s.ToString();
                    if (!string.IsNullOrWhiteSpace(val))
                        result.SpecialRequirements.Add(val);
                }
            }

            return result;
        }

        /// <summary>
        /// Safely parses JSON, stripping markdown fencing if present.
        /// </summary>
        private static JObject SafeParseJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return null;

            // Strip markdown code fencing if the AI added it
            json = json.Trim();
            if (json.StartsWith("```json", StringComparison.OrdinalIgnoreCase))
                json = json.Substring(7);
            else if (json.StartsWith("```"))
                json = json.Substring(3);
            if (json.EndsWith("```"))
                json = json.Substring(0, json.Length - 3);
            json = json.Trim();

            try
            {
                return JObject.Parse(json);
            }
            catch (Exception)
            {
                // Try to find JSON object within the text
                int start = json.IndexOf('{');
                int end = json.LastIndexOf('}');
                if (start >= 0 && end > start)
                {
                    try
                    {
                        return JObject.Parse(json.Substring(start, end - start + 1));
                    }
                    catch
                    {
                        return null;
                    }
                }
                return null;
            }
        }

        private static FieldResult ExtractField(JObject obj, string key, double defaultConfidence = 0.9)
        {
            var token = obj[key];
            if (token == null || token.Type == JTokenType.Null)
                return new FieldResult("", 0);

            string value = token.ToString().Trim();
            if (string.IsNullOrEmpty(value))
                return new FieldResult("", 0);

            return new FieldResult(value, defaultConfidence);
        }

        private static FieldResult ExtractNumericField(JObject obj, string key)
        {
            var token = obj[key];
            if (token == null || token.Type == JTokenType.Null)
                return null;

            if (token.Type == JTokenType.Float || token.Type == JTokenType.Integer)
            {
                double val = token.Value<double>();
                return new FieldResult(val.ToString("F4", CultureInfo.InvariantCulture), 0.85);
            }

            // Try parsing string value
            string strVal = token.ToString().Trim();
            if (double.TryParse(strVal, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed))
            {
                return new FieldResult(parsed.ToString("F4", CultureInfo.InvariantCulture), 0.80);
            }

            return null;
        }

        private static void ExtractStringArray(JObject obj, string key, List<string> target)
        {
            var arr = obj[key] as JArray;
            if (arr == null) return;

            foreach (var item in arr)
            {
                string val = item.ToString().Trim();
                if (!string.IsNullOrWhiteSpace(val))
                    target.Add(val);
            }
        }
    }
}
