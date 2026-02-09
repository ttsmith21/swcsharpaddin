using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using NM.Core.Pdf.Models;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Measures extraction accuracy by comparing DrawingData against ground truth.
    /// Reports precision, recall, and F1 score per field.
    /// </summary>
    public sealed class AccuracyBenchmark
    {
        /// <summary>
        /// Runs extraction on all ground truth files in a directory and reports accuracy.
        /// </summary>
        /// <param name="testDir">Directory containing PDF files and their .gt.json ground truth files.</param>
        /// <param name="analyzer">The analyzer to test.</param>
        /// <returns>Benchmark results with per-field accuracy metrics.</returns>
        public BenchmarkReport Run(string testDir, PdfDrawingAnalyzer analyzer)
        {
            var report = new BenchmarkReport();

            if (!Directory.Exists(testDir))
            {
                report.Error = $"Test directory not found: {testDir}";
                return report;
            }

            // Find all ground truth JSON files
            var gtFiles = Directory.GetFiles(testDir, "*.gt.json", SearchOption.TopDirectoryOnly);
            if (gtFiles.Length == 0)
            {
                report.Error = "No ground truth files (*.gt.json) found in test directory";
                return report;
            }

            foreach (var gtFile in gtFiles)
            {
                try
                {
                    string json = File.ReadAllText(gtFile);
                    var gt = JsonConvert.DeserializeObject<ExtractionGroundTruth>(json);
                    if (gt == null) continue;

                    string pdfPath = Path.Combine(testDir, gt.PdfFileName);
                    if (!File.Exists(pdfPath))
                    {
                        report.Skipped.Add(gt.PdfFileName, "PDF file not found");
                        continue;
                    }

                    var extracted = analyzer.Analyze(pdfPath);
                    var result = CompareResult(gt, extracted);
                    report.Results.Add(result);
                }
                catch (Exception ex)
                {
                    report.Skipped.Add(Path.GetFileName(gtFile), ex.Message);
                }
            }

            report.CalculateSummary();
            return report;
        }

        /// <summary>
        /// Compares a single extraction result against ground truth.
        /// </summary>
        public DrawingComparisonResult CompareResult(ExtractionGroundTruth gt, DrawingData extracted)
        {
            var result = new DrawingComparisonResult
            {
                PdfFileName = gt.PdfFileName,
                DrawingType = gt.DrawingType
            };

            // Title block fields
            result.Fields.Add(CompareField("PartNumber", gt.PartNumber, extracted.PartNumber));
            result.Fields.Add(CompareField("Description", gt.Description, extracted.Description));
            result.Fields.Add(CompareField("Revision", gt.Revision, extracted.Revision));
            result.Fields.Add(CompareField("Material", gt.Material, extracted.Material));
            result.Fields.Add(CompareField("Finish", gt.Finish, extracted.Finish));
            result.Fields.Add(CompareField("DrawnBy", gt.DrawnBy, extracted.DrawnBy));
            result.Fields.Add(CompareField("Scale", gt.Scale, extracted.Scale));
            result.Fields.Add(CompareField("SheetInfo", gt.SheetInfo, extracted.SheetInfo));
            result.Fields.Add(CompareField("ToleranceGeneral", gt.ToleranceGeneral, extracted.ToleranceGeneral));

            // Manufacturing notes: check which ground truth notes were found
            var extractedNoteTexts = extracted.Notes.Select(n => n.Text).ToList();
            result.NotesMetrics = CompareNoteLists(gt.ManufacturingNotes, extractedNoteTexts);

            return result;
        }

        private FieldComparison CompareField(string fieldName, string expected, string actual)
        {
            bool expectedEmpty = string.IsNullOrWhiteSpace(expected);
            bool actualEmpty = string.IsNullOrWhiteSpace(actual);

            if (expectedEmpty && actualEmpty)
            {
                return new FieldComparison
                {
                    FieldName = fieldName,
                    Expected = "",
                    Actual = "",
                    Match = MatchType.TrueNegative
                };
            }

            if (expectedEmpty && !actualEmpty)
            {
                return new FieldComparison
                {
                    FieldName = fieldName,
                    Expected = "",
                    Actual = actual,
                    Match = MatchType.FalsePositive
                };
            }

            if (!expectedEmpty && actualEmpty)
            {
                return new FieldComparison
                {
                    FieldName = fieldName,
                    Expected = expected,
                    Actual = "",
                    Match = MatchType.FalseNegative
                };
            }

            // Both have values — check if they match
            bool exactMatch = string.Equals(expected.Trim(), actual.Trim(), StringComparison.OrdinalIgnoreCase);
            bool fuzzyMatch = !exactMatch && (
                expected.IndexOf(actual, StringComparison.OrdinalIgnoreCase) >= 0 ||
                actual.IndexOf(expected, StringComparison.OrdinalIgnoreCase) >= 0);

            return new FieldComparison
            {
                FieldName = fieldName,
                Expected = expected,
                Actual = actual,
                Match = exactMatch ? MatchType.TruePositiveExact
                    : fuzzyMatch ? MatchType.TruePositiveFuzzy
                    : MatchType.Mismatch
            };
        }

        private NoteListMetrics CompareNoteLists(List<string> expected, List<string> actual)
        {
            var metrics = new NoteListMetrics();
            if (expected == null) expected = new List<string>();
            if (actual == null) actual = new List<string>();

            var expectedSet = new HashSet<string>(expected, StringComparer.OrdinalIgnoreCase);
            var actualSet = new HashSet<string>(actual, StringComparer.OrdinalIgnoreCase);

            // True positives: notes in both lists (fuzzy matching)
            foreach (var exp in expected)
            {
                bool found = actual.Any(a =>
                    a.IndexOf(exp, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    exp.IndexOf(a, StringComparison.OrdinalIgnoreCase) >= 0);

                if (found)
                    metrics.TruePositives++;
                else
                    metrics.FalseNegatives++;
            }

            // False positives: notes extracted but not in ground truth
            foreach (var act in actual)
            {
                bool found = expected.Any(e =>
                    e.IndexOf(act, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    act.IndexOf(e, StringComparison.OrdinalIgnoreCase) >= 0);

                if (!found)
                    metrics.FalsePositives++;
            }

            metrics.ExpectedCount = expected.Count;
            metrics.ActualCount = actual.Count;

            return metrics;
        }
    }

    /// <summary>
    /// Full benchmark report across multiple test drawings.
    /// </summary>
    public sealed class BenchmarkReport
    {
        public List<DrawingComparisonResult> Results { get; } = new List<DrawingComparisonResult>();
        public Dictionary<string, string> Skipped { get; } = new Dictionary<string, string>();
        public string Error { get; set; }

        // Summary metrics (calculated after all results are added)
        public Dictionary<string, FieldMetrics> FieldSummary { get; } = new Dictionary<string, FieldMetrics>();
        public NoteListMetrics NotesSummary { get; set; }

        public void CalculateSummary()
        {
            if (Results.Count == 0) return;

            // Aggregate field-level metrics
            var allFieldNames = Results.SelectMany(r => r.Fields.Select(f => f.FieldName)).Distinct();
            foreach (var fieldName in allFieldNames)
            {
                var comparisons = Results.SelectMany(r => r.Fields.Where(f => f.FieldName == fieldName)).ToList();
                var metrics = new FieldMetrics
                {
                    TruePositives = comparisons.Count(c => c.Match == MatchType.TruePositiveExact || c.Match == MatchType.TruePositiveFuzzy),
                    FalsePositives = comparisons.Count(c => c.Match == MatchType.FalsePositive || c.Match == MatchType.Mismatch),
                    FalseNegatives = comparisons.Count(c => c.Match == MatchType.FalseNegative),
                    TrueNegatives = comparisons.Count(c => c.Match == MatchType.TrueNegative),
                    Total = comparisons.Count
                };
                FieldSummary[fieldName] = metrics;
            }

            // Aggregate note-level metrics
            NotesSummary = new NoteListMetrics
            {
                TruePositives = Results.Sum(r => r.NotesMetrics?.TruePositives ?? 0),
                FalsePositives = Results.Sum(r => r.NotesMetrics?.FalsePositives ?? 0),
                FalseNegatives = Results.Sum(r => r.NotesMetrics?.FalseNegatives ?? 0),
                ExpectedCount = Results.Sum(r => r.NotesMetrics?.ExpectedCount ?? 0),
                ActualCount = Results.Sum(r => r.NotesMetrics?.ActualCount ?? 0)
            };
        }

        /// <summary>
        /// Formats the report as a human-readable string.
        /// </summary>
        public string FormatReport()
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("=== PDF Extraction Accuracy Benchmark ===");
            sb.AppendLine($"Drawings tested: {Results.Count}, Skipped: {Skipped.Count}");
            sb.AppendLine();

            if (FieldSummary.Count > 0)
            {
                sb.AppendLine("--- Per-Field Metrics ---");
                sb.AppendLine($"{"Field",-20} {"Precision",10} {"Recall",10} {"F1",10} {"TP",5} {"FP",5} {"FN",5}");
                foreach (var kvp in FieldSummary)
                {
                    var m = kvp.Value;
                    sb.AppendLine($"{kvp.Key,-20} {m.Precision,10:P1} {m.Recall,10:P1} {m.F1,10:P1} {m.TruePositives,5} {m.FalsePositives,5} {m.FalseNegatives,5}");
                }
            }

            if (NotesSummary != null)
            {
                sb.AppendLine();
                sb.AppendLine("--- Notes Metrics ---");
                sb.AppendLine($"Precision: {NotesSummary.Precision:P1}, Recall: {NotesSummary.Recall:P1}, F1: {NotesSummary.F1:P1}");
                sb.AppendLine($"TP: {NotesSummary.TruePositives}, FP: {NotesSummary.FalsePositives}, FN: {NotesSummary.FalseNegatives}");
            }

            return sb.ToString();
        }
    }

    public sealed class DrawingComparisonResult
    {
        public string PdfFileName { get; set; }
        public string DrawingType { get; set; }
        public List<FieldComparison> Fields { get; } = new List<FieldComparison>();
        public NoteListMetrics NotesMetrics { get; set; }
    }

    public sealed class FieldComparison
    {
        public string FieldName { get; set; }
        public string Expected { get; set; }
        public string Actual { get; set; }
        public MatchType Match { get; set; }
    }

    public enum MatchType
    {
        TruePositiveExact,
        TruePositiveFuzzy,
        FalsePositive,
        FalseNegative,
        TrueNegative,
        Mismatch
    }

    public sealed class FieldMetrics
    {
        public int TruePositives { get; set; }
        public int FalsePositives { get; set; }
        public int FalseNegatives { get; set; }
        public int TrueNegatives { get; set; }
        public int Total { get; set; }

        public double Precision => (TruePositives + FalsePositives) > 0
            ? (double)TruePositives / (TruePositives + FalsePositives) : 1.0;
        public double Recall => (TruePositives + FalseNegatives) > 0
            ? (double)TruePositives / (TruePositives + FalseNegatives) : 1.0;
        public double F1
        {
            get
            {
                double p = Precision, r = Recall;
                return (p + r) > 0 ? 2 * p * r / (p + r) : 0;
            }
        }
    }

    public sealed class NoteListMetrics
    {
        public int TruePositives { get; set; }
        public int FalsePositives { get; set; }
        public int FalseNegatives { get; set; }
        public int ExpectedCount { get; set; }
        public int ActualCount { get; set; }

        public double Precision => (TruePositives + FalsePositives) > 0
            ? (double)TruePositives / (TruePositives + FalsePositives) : 1.0;
        public double Recall => (TruePositives + FalseNegatives) > 0
            ? (double)TruePositives / (TruePositives + FalseNegatives) : 1.0;
        public double F1
        {
            get
            {
                double p = Precision, r = Recall;
                return (p + r) > 0 ? 2 * p * r / (p + r) : 0;
            }
        }
    }
}
