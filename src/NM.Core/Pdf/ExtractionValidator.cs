using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using NM.Core.Pdf.Models;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Validates extracted drawing data against domain knowledge to catch false positives.
    /// Checks material names, finish types, part number formats, and impossible values.
    /// </summary>
    public sealed class ExtractionValidator
    {
        // Known material families (case-insensitive matching)
        private static readonly HashSet<string> KnownMaterialKeywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "STEEL", "STAINLESS", "ALUMINUM", "ALUMINIUM", "COPPER", "BRASS", "BRONZE",
            "TITANIUM", "INCONEL", "MONEL", "HASTELLOY", "NICKEL",
            "A36", "A53", "A500", "A513", "A514", "A572",
            "304", "304L", "316", "316L", "321", "347", "410", "430", "440",
            "1008", "1010", "1018", "1020", "1045", "1095",
            "4130", "4140", "4340", "8620",
            "6061", "5052", "3003", "2024", "7075",
            "CRS", "HRS", "HRPO", "CR", "HR",
            "SS", "CS", "MS", "AL",
            "GALVANIZED", "GALVANNEAL", "GALV",
            "ASTM", "SAE", "AISI", "AMS", "MIL",
            "MILD STEEL", "CARBON STEEL", "STAINLESS STEEL",
            "DOM", "ERW", "SEAMLESS"
        };

        // Known finish types
        private static readonly HashSet<string> KnownFinishKeywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "PAINT", "POWDER COAT", "ANODIZE", "GALVANIZE", "ZINC PLATE",
            "CHROME PLATE", "BLACK OXIDE", "E-COAT", "PRIME", "PRIMER",
            "HOT DIP", "ELECTROLESS NICKEL", "HARD CHROME", "PASSIVATE",
            "CHEM FILM", "ALODINE", "CONVERSION COATING", "CLEAR COAT",
            "NONE", "N/A", "AS MACHINED", "MILL FINISH",
            "SANDBLAST", "BEAD BLAST", "TUMBLE",
            "POLISHED", "BRUSHED", "SATIN"
        };

        // Words that should never appear in a valid part number
        private static readonly string[] InvalidPartNumberWords =
        {
            "SCALE", "MATERIAL", "FINISH", "DRAWN", "CHECKED", "DATE",
            "REVISION", "SHEET", "TITLE", "DESCRIPTION", "UNLESS",
            "TOLERANCE", "DO NOT", "BREAK", "DEBURR", "PAINT",
            "NOTES", "GENERAL", "DIMENSIONS", "SPECIFIED"
        };

        /// <summary>
        /// Validates extracted DrawingData and returns a list of issues found.
        /// </summary>
        public List<ValidationIssue> Validate(DrawingData data)
        {
            var issues = new List<ValidationIssue>();
            if (data == null) return issues;

            ValidatePartNumber(data.PartNumber, issues);
            ValidateMaterial(data.Material, issues);
            ValidateFinish(data.Finish, issues);
            ValidateThickness(data.Thickness_in, issues);
            ValidateDimensions(data.OverallLength_in, data.OverallWidth_in, issues);
            ValidateNotes(data.Notes, issues);

            return issues;
        }

        /// <summary>
        /// Applies validation to fix or flag confidence on extracted data.
        /// Returns the number of fields corrected.
        /// </summary>
        public int ValidateAndCorrect(DrawingData data)
        {
            var issues = Validate(data);
            int corrections = 0;

            foreach (var issue in issues)
            {
                if (issue.Severity == IssueSeverity.Error)
                {
                    // Clear obviously wrong fields
                    switch (issue.Field)
                    {
                        case "PartNumber":
                            data.PartNumber = null;
                            corrections++;
                            break;
                        case "Material":
                            data.Material = null;
                            corrections++;
                            break;
                        case "Finish":
                            data.Finish = null;
                            corrections++;
                            break;
                        case "Thickness":
                            data.Thickness_in = null;
                            corrections++;
                            break;
                    }
                }
            }

            // Remove notes that failed validation
            var invalidNotes = issues
                .Where(i => i.Field.StartsWith("Note[") && i.Severity == IssueSeverity.Error)
                .Select(i => i.OriginalValue)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (invalidNotes.Count > 0)
            {
                data.Notes.RemoveAll(n => invalidNotes.Contains(n.Text));
                corrections += invalidNotes.Count;
            }

            return corrections;
        }

        private void ValidatePartNumber(string partNumber, List<ValidationIssue> issues)
        {
            if (string.IsNullOrEmpty(partNumber)) return;

            // Check for title block label leakage
            foreach (var word in InvalidPartNumberWords)
            {
                if (partNumber.IndexOf(word, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    issues.Add(new ValidationIssue
                    {
                        Field = "PartNumber",
                        OriginalValue = partNumber,
                        Severity = IssueSeverity.Error,
                        Message = $"Part number contains label text '{word}' — likely a parsing error"
                    });
                    return;
                }
            }

            // Part numbers should be mostly alphanumeric with dashes/dots
            if (!Regex.IsMatch(partNumber, @"^[A-Za-z0-9][\w\-\.\/]{1,30}$"))
            {
                issues.Add(new ValidationIssue
                {
                    Field = "PartNumber",
                    OriginalValue = partNumber,
                    Severity = IssueSeverity.Warning,
                    Message = "Part number has unusual format"
                });
            }
        }

        private void ValidateMaterial(string material, List<ValidationIssue> issues)
        {
            if (string.IsNullOrEmpty(material)) return;

            // Check if any known material keyword is present
            bool hasKnownKeyword = KnownMaterialKeywords.Any(k =>
                material.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0);

            if (!hasKnownKeyword)
            {
                issues.Add(new ValidationIssue
                {
                    Field = "Material",
                    OriginalValue = material,
                    Severity = IssueSeverity.Warning,
                    Message = "Material does not match any known material keyword"
                });
            }

            // Check for label leakage (common false positive)
            if (Regex.IsMatch(material, @"\b(FINISH|SCALE|DRAWN|DATE|REV|SHEET)\b", RegexOptions.IgnoreCase))
            {
                issues.Add(new ValidationIssue
                {
                    Field = "Material",
                    OriginalValue = material,
                    Severity = IssueSeverity.Error,
                    Message = "Material field contains adjacent title block label — likely a parsing error"
                });
            }
        }

        private void ValidateFinish(string finish, List<ValidationIssue> issues)
        {
            if (string.IsNullOrEmpty(finish)) return;

            bool hasKnownKeyword = KnownFinishKeywords.Any(k =>
                finish.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0);

            if (!hasKnownKeyword)
            {
                issues.Add(new ValidationIssue
                {
                    Field = "Finish",
                    OriginalValue = finish,
                    Severity = IssueSeverity.Warning,
                    Message = "Finish does not match any known finish type"
                });
            }
        }

        private void ValidateThickness(double? thickness, List<ValidationIssue> issues)
        {
            if (!thickness.HasValue) return;

            if (thickness.Value <= 0)
            {
                issues.Add(new ValidationIssue
                {
                    Field = "Thickness",
                    OriginalValue = thickness.Value.ToString("F4"),
                    Severity = IssueSeverity.Error,
                    Message = "Thickness must be positive"
                });
            }
            else if (thickness.Value > 12.0)
            {
                issues.Add(new ValidationIssue
                {
                    Field = "Thickness",
                    OriginalValue = thickness.Value.ToString("F4"),
                    Severity = IssueSeverity.Warning,
                    Message = "Thickness exceeds 12 inches — verify units"
                });
            }
        }

        private void ValidateDimensions(double? length, double? width, List<ValidationIssue> issues)
        {
            if (length.HasValue && length.Value <= 0)
            {
                issues.Add(new ValidationIssue
                {
                    Field = "OverallLength",
                    OriginalValue = length.Value.ToString("F4"),
                    Severity = IssueSeverity.Error,
                    Message = "Length must be positive"
                });
            }
            if (width.HasValue && width.Value <= 0)
            {
                issues.Add(new ValidationIssue
                {
                    Field = "OverallWidth",
                    OriginalValue = width.Value.ToString("F4"),
                    Severity = IssueSeverity.Error,
                    Message = "Width must be positive"
                });
            }
        }

        private void ValidateNotes(List<DrawingNote> notes, List<ValidationIssue> issues)
        {
            if (notes == null) return;

            for (int i = 0; i < notes.Count; i++)
            {
                var note = notes[i];

                // Notes that are too short are suspicious
                if (note.Text != null && note.Text.Length < 3)
                {
                    issues.Add(new ValidationIssue
                    {
                        Field = $"Note[{i}]",
                        OriginalValue = note.Text,
                        Severity = IssueSeverity.Warning,
                        Message = "Note text is suspiciously short"
                    });
                }

                // Notes that are just numbers are likely BOM entries, not manufacturing notes
                if (note.Text != null && Regex.IsMatch(note.Text.Trim(), @"^\d+$"))
                {
                    issues.Add(new ValidationIssue
                    {
                        Field = $"Note[{i}]",
                        OriginalValue = note.Text,
                        Severity = IssueSeverity.Error,
                        Message = "Note is just a number — likely a BOM entry, not a manufacturing note"
                    });
                }
            }
        }
    }

    /// <summary>
    /// A validation issue found in extracted drawing data.
    /// </summary>
    public sealed class ValidationIssue
    {
        public string Field { get; set; }
        public string OriginalValue { get; set; }
        public IssueSeverity Severity { get; set; }
        public string Message { get; set; }

        public override string ToString() => $"[{Severity}] {Field}: {Message} (was: '{OriginalValue}')";
    }

    public enum IssueSeverity
    {
        Info,
        Warning,
        Error
    }
}
