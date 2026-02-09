using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Data-driven confidence calibration based on empirical accuracy measurements.
    /// Replaces hardcoded confidence scores with values derived from the test corpus.
    /// Calibration data is stored as JSON and updated as the corpus grows.
    /// </summary>
    public sealed class ConfidenceCalibrator
    {
        private readonly Dictionary<string, double> _patternPrecision;
        private readonly Dictionary<string, double> _fieldAccuracy;

        public ConfidenceCalibrator()
        {
            _patternPrecision = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            _fieldAccuracy = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Loads calibration data from a JSON file.
        /// </summary>
        public static ConfidenceCalibrator LoadFromFile(string path)
        {
            var calibrator = new ConfidenceCalibrator();
            if (!File.Exists(path)) return calibrator;

            try
            {
                var data = JsonConvert.DeserializeObject<CalibrationData>(File.ReadAllText(path));
                if (data?.PatternPrecision != null)
                {
                    foreach (var kvp in data.PatternPrecision)
                        calibrator._patternPrecision[kvp.Key] = kvp.Value;
                }
                if (data?.FieldAccuracy != null)
                {
                    foreach (var kvp in data.FieldAccuracy)
                        calibrator._fieldAccuracy[kvp.Key] = kvp.Value;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Calibrator] Failed to load calibration data: {ex.Message}");
            }

            return calibrator;
        }

        /// <summary>
        /// Saves calibration data to a JSON file.
        /// </summary>
        public void SaveToFile(string path)
        {
            var data = new CalibrationData
            {
                PatternPrecision = new Dictionary<string, double>(_patternPrecision),
                FieldAccuracy = new Dictionary<string, double>(_fieldAccuracy),
                LastUpdated = DateTime.UtcNow
            };

            string json = JsonConvert.SerializeObject(data, Formatting.Indented);
            File.WriteAllText(path, json);
        }

        /// <summary>
        /// Gets the calibrated confidence for a regex pattern.
        /// Falls back to the default confidence if no calibration data exists.
        /// </summary>
        public double GetPatternConfidence(string patternTemplate, double defaultConfidence)
        {
            if (_patternPrecision.TryGetValue(patternTemplate, out double calibrated))
                return calibrated;
            return defaultConfidence;
        }

        /// <summary>
        /// Gets the calibrated accuracy for a title block field.
        /// </summary>
        public double GetFieldConfidence(string fieldName, double defaultConfidence)
        {
            if (_fieldAccuracy.TryGetValue(fieldName, out double calibrated))
                return calibrated;
            return defaultConfidence;
        }

        /// <summary>
        /// Updates calibration data from benchmark results.
        /// Call this after running AccuracyBenchmark to recalibrate.
        /// </summary>
        public void UpdateFromBenchmark(BenchmarkReport report)
        {
            if (report == null) return;

            // Update field accuracy from benchmark field summary
            foreach (var kvp in report.FieldSummary)
            {
                _fieldAccuracy[kvp.Key] = kvp.Value.Precision;
            }
        }

        /// <summary>
        /// Records a pattern match result for ongoing calibration.
        /// </summary>
        public void RecordPatternResult(string patternTemplate, bool wasCorrect)
        {
            string trueKey = patternTemplate + "_true";
            string totalKey = patternTemplate + "_total";

            if (!_patternPrecision.ContainsKey(totalKey))
            {
                _patternPrecision[totalKey] = 0;
                _patternPrecision[trueKey] = 0;
            }

            _patternPrecision[totalKey]++;
            if (wasCorrect) _patternPrecision[trueKey]++;

            // Recalculate precision
            double total = _patternPrecision[totalKey];
            double trueCount = _patternPrecision[trueKey];
            _patternPrecision[patternTemplate] = total > 0 ? trueCount / total : 0.5;
        }

        /// <summary>
        /// Adjusts a confidence score based on cross-validation between text and vision sources.
        /// Both sources agreeing boosts confidence; disagreement reduces it.
        /// </summary>
        public static double CrossValidateConfidence(
            double textConfidence, bool textFound,
            double visionConfidence, bool visionFound)
        {
            if (textFound && visionFound)
            {
                // Both sources agree — boost confidence
                return Math.Min(1.0, Math.Max(textConfidence, visionConfidence) * 1.15);
            }
            else if (textFound && !visionFound)
            {
                // Only text found it — reduce confidence slightly
                return textConfidence * 0.85;
            }
            else if (!textFound && visionFound)
            {
                // Only vision found it — reduce confidence slightly
                return visionConfidence * 0.85;
            }
            else
            {
                return 0;
            }
        }

        /// <summary>
        /// Checks if a drawing has suspiciously low extraction density.
        /// Multi-page drawings with very few notes may indicate false negatives.
        /// </summary>
        public static CoverageDensityResult CheckCoverageDensity(
            int pageCount, int noteCount, int gdtCount, bool hasTolerances, bool hasTitleBlock = false)
        {
            var result = new CoverageDensityResult();

            if (pageCount >= 3 && noteCount <= 1)
            {
                result.Suspicious = true;
                result.Reason = $"Multi-page drawing ({pageCount} pages) has only {noteCount} note(s) — possible false negatives";
            }

            if (pageCount >= 2 && noteCount == 0 && gdtCount == 0)
            {
                result.Suspicious = true;
                result.Reason = $"Multi-page drawing ({pageCount} pages) has zero notes and zero GD&T — likely extraction failure";
            }

            // Title block populated but zero notes — may indicate notes section was missed
            if (hasTitleBlock && noteCount == 0 && pageCount >= 1)
            {
                result.Suspicious = true;
                result.Reason = "Title block populated but zero manufacturing notes extracted — notes section may have been missed";
            }

            return result;
        }
    }

    /// <summary>
    /// Serializable calibration data.
    /// </summary>
    internal sealed class CalibrationData
    {
        public Dictionary<string, double> PatternPrecision { get; set; }
        public Dictionary<string, double> FieldAccuracy { get; set; }
        public DateTime LastUpdated { get; set; }
    }

    /// <summary>
    /// Result of coverage density check.
    /// </summary>
    public sealed class CoverageDensityResult
    {
        public bool Suspicious { get; set; }
        public string Reason { get; set; }
    }
}
