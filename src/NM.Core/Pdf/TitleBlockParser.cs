using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using NM.Core.Pdf.Models;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Parses title block data from extracted PDF text using regex patterns.
    /// Handles common engineering drawing title block formats.
    /// </summary>
    public sealed class TitleBlockParser
    {
        // --- Part Number patterns ---
        private static readonly Regex[] PartNumberPatterns = new[]
        {
            new Regex(@"(?:PART\s*(?:NO|NUMBER|#|NUM)\.?\s*[:.]?\s*)([A-Z0-9][\w\-\.]+)", RegexOptions.IgnoreCase),
            new Regex(@"(?:DWG\s*(?:NO|NUMBER|#|NUM)\.?\s*[:.]?\s*)([A-Z0-9][\w\-\.]+)", RegexOptions.IgnoreCase),
            new Regex(@"(?:DRAWING\s*(?:NO|NUMBER|#|NUM)\.?\s*[:.]?\s*)([A-Z0-9][\w\-\.]+)", RegexOptions.IgnoreCase),
            new Regex(@"(?:P/?N\s*[:.]?\s*)([A-Z0-9][\w\-\.]+)", RegexOptions.IgnoreCase),
            new Regex(@"(?:ITEM\s*(?:NO|NUMBER|#)\.?\s*[:.]?\s*)([A-Z0-9][\w\-\.]+)", RegexOptions.IgnoreCase),
        };

        // --- Material patterns ---
        private static readonly Regex[] MaterialPatterns = new[]
        {
            new Regex(@"(?:MATERIAL\s*[:.]?\s*)(.+?)(?:\s*$|\s*FINISH|\s*SCALE|\s*UNLESS)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
            new Regex(@"(?:MAT(?:'?L)?\s*[:.]?\s*)(.+?)(?:\s*$|\s*FINISH|\s*SCALE)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
            new Regex(@"(?:MATL\s*SPEC\s*[:.]?\s*)(.+?)(?:\s*$)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
        };

        // --- Specific material callouts (standalone matches) ---
        private static readonly Regex[] MaterialCalloutPatterns = new[]
        {
            new Regex(@"\b(ASTM\s*A[\-\s]?\d+)", RegexOptions.IgnoreCase),
            new Regex(@"\b(A36|A53[12]?|A500)\b", RegexOptions.IgnoreCase),
            new Regex(@"\b(\d{3,4}\s*(?:STAINLESS|SS|CRS|HRS|AL|ALUMINUM))\b", RegexOptions.IgnoreCase),
            new Regex(@"\b(304L?|316L?|1018|1020|1045|4130|4140|6061|5052|3003)\s*(?:SS|STAINLESS|CRS|HRS|AL|ALUMINUM|STEEL)?\b", RegexOptions.IgnoreCase),
            new Regex(@"\b(MILD\s*STEEL|CARBON\s*STEEL|STAINLESS\s*STEEL|GALVANIZED|GALVANNEAL)\b", RegexOptions.IgnoreCase),
        };

        // --- Revision patterns ---
        private static readonly Regex[] RevisionPatterns = new[]
        {
            new Regex(@"(?:REV(?:ISION)?\.?\s*[:.]?\s*)([A-Z0-9]{1,3})\b", RegexOptions.IgnoreCase),
            new Regex(@"\bREV\s+([A-Z])\b", RegexOptions.IgnoreCase),
        };

        // --- Description patterns ---
        private static readonly Regex[] DescriptionPatterns = new[]
        {
            new Regex(@"(?:DESC(?:RIPTION)?\.?\s*[:.]?\s*)(.+?)(?:\s*$)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
            new Regex(@"(?:TITLE\s*[:.]?\s*)(.+?)(?:\s*$)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
            new Regex(@"(?:NAME\s*[:.]?\s*)(.+?)(?:\s*$)", RegexOptions.IgnoreCase | RegexOptions.Multiline),
        };

        // --- Other fields ---
        private static readonly Regex DrawnByPattern = new Regex(@"(?:DRAWN\s*(?:BY)?\s*[:.]?\s*)([A-Z][A-Z\s\.]{1,20})", RegexOptions.IgnoreCase);
        private static readonly Regex CheckedByPattern = new Regex(@"(?:CHE?C?KE?D?\s*(?:BY)?\s*[:.]?\s*)([A-Z][A-Z\s\.]{1,20})", RegexOptions.IgnoreCase);
        private static readonly Regex ScalePattern = new Regex(@"(?:SCALE\s*[:.]?\s*)([\d]+\s*[:/]\s*[\d]+|FULL|HALF|NTS|NONE)", RegexOptions.IgnoreCase);
        private static readonly Regex SheetPattern = new Regex(@"(?:SHEET\s*[:.]?\s*)(\d+\s*(?:OF|/)\s*\d+)", RegexOptions.IgnoreCase);
        private static readonly Regex DatePattern = new Regex(@"(?:DATE\s*[:.]?\s*)(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})", RegexOptions.IgnoreCase);
        private static readonly Regex TolerancePattern = new Regex(@"(?:UNLESS\s+OTHERWISE\s+(?:NOTED|SPECIFIED|STATED).*?TOLERANCES?\s*[:.]?\s*)(.+?)(?:\s*$)", RegexOptions.IgnoreCase | RegexOptions.Singleline);

        /// <summary>
        /// Parses title block information from raw text.
        /// Tries all patterns and returns the best match for each field.
        /// </summary>
        public TitleBlockInfo Parse(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new TitleBlockInfo();

            var info = new TitleBlockInfo();

            // Part number
            var pn = TryMatch(text, PartNumberPatterns);
            if (pn != null)
            {
                info.PartNumber = CleanValue(pn);
                info.PartNumberConfidence = 0.85;
            }

            // Material — try labeled patterns first, then standalone callouts
            var mat = TryMatch(text, MaterialPatterns);
            if (mat != null)
            {
                info.Material = CleanMaterial(mat);
                info.MaterialConfidence = 0.85;
            }
            else
            {
                mat = TryMatch(text, MaterialCalloutPatterns);
                if (mat != null)
                {
                    info.Material = CleanMaterial(mat);
                    info.MaterialConfidence = 0.70; // Lower confidence for standalone callouts
                }
            }

            // Revision
            var rev = TryMatch(text, RevisionPatterns);
            if (rev != null)
            {
                info.Revision = rev.Trim().ToUpperInvariant();
                info.RevisionConfidence = 0.90;
            }

            // Description
            var desc = TryMatch(text, DescriptionPatterns);
            if (desc != null)
            {
                info.Description = CleanValue(desc);
                info.DescriptionConfidence = 0.80;
            }

            // Simple single-pattern fields
            info.DrawnBy = MatchSingle(text, DrawnByPattern);
            info.CheckedBy = MatchSingle(text, CheckedByPattern);
            info.Scale = MatchSingle(text, ScalePattern);
            info.Sheet = MatchSingle(text, SheetPattern);
            info.ToleranceGeneral = MatchSingle(text, TolerancePattern);

            // Date parsing
            var dateStr = MatchSingle(text, DatePattern);
            if (dateStr != null)
            {
                if (DateTime.TryParse(dateStr, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                    info.Date = dt;
            }

            return info;
        }

        /// <summary>
        /// Parses title block from the full text of a page, focusing on the
        /// bottom-right region where title blocks typically appear.
        /// </summary>
        public TitleBlockInfo ParseFromPage(PageText page, PdfTextExtractor extractor)
        {
            if (page == null) return new TitleBlockInfo();

            // Try title block region first
            string titleBlockText = extractor.ExtractTitleBlockRegion(page);
            var result = Parse(titleBlockText);

            // If we didn't get key fields, try the full page text
            if (string.IsNullOrEmpty(result.PartNumber) || string.IsNullOrEmpty(result.Material))
            {
                var fullResult = Parse(page.FullText);

                if (string.IsNullOrEmpty(result.PartNumber) && !string.IsNullOrEmpty(fullResult.PartNumber))
                {
                    result.PartNumber = fullResult.PartNumber;
                    result.PartNumberConfidence = fullResult.PartNumberConfidence * 0.9; // Slightly lower
                }
                if (string.IsNullOrEmpty(result.Material) && !string.IsNullOrEmpty(fullResult.Material))
                {
                    result.Material = fullResult.Material;
                    result.MaterialConfidence = fullResult.MaterialConfidence * 0.9;
                }
                if (string.IsNullOrEmpty(result.Revision) && !string.IsNullOrEmpty(fullResult.Revision))
                {
                    result.Revision = fullResult.Revision;
                    result.RevisionConfidence = fullResult.RevisionConfidence * 0.9;
                }
                if (string.IsNullOrEmpty(result.Description) && !string.IsNullOrEmpty(fullResult.Description))
                {
                    result.Description = fullResult.Description;
                    result.DescriptionConfidence = fullResult.DescriptionConfidence * 0.9;
                }
            }

            return result;
        }

        private static string TryMatch(string text, Regex[] patterns)
        {
            foreach (var pattern in patterns)
            {
                var match = pattern.Match(text);
                if (match.Success && match.Groups.Count > 1)
                {
                    string value = match.Groups[1].Value.Trim();
                    if (!string.IsNullOrWhiteSpace(value))
                        return value;
                }
            }
            return null;
        }

        private static string MatchSingle(string text, Regex pattern)
        {
            var match = pattern.Match(text);
            if (match.Success && match.Groups.Count > 1)
            {
                string value = match.Groups[1].Value.Trim();
                return string.IsNullOrWhiteSpace(value) ? null : value;
            }
            return null;
        }

        private static string CleanValue(string value)
        {
            if (value == null) return null;
            return value.Trim().Trim(':', '.', ',', '-');
        }

        private static string CleanMaterial(string value)
        {
            if (value == null) return null;
            value = value.Trim().Trim(':', '.', ',');
            // Remove trailing labels that leaked in
            value = Regex.Replace(value, @"\s*(FINISH|SCALE|UNLESS|DRAWN|DATE|REV).*$", "", RegexOptions.IgnoreCase);
            return value.Trim();
        }
    }
}
