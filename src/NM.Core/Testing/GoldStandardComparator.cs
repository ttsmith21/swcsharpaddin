using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace NM.Core.Testing
{
    /// <summary>
    /// Compares QA runner results against gold standard manifest
    /// </summary>
    public sealed class GoldStandardComparator
    {
        /// <summary>
        /// Compare results.json against manifest-v2.json
        /// </summary>
        public static FullComparisonReport Compare(string manifestPath, string resultsPath)
        {
            if (!File.Exists(manifestPath))
                throw new FileNotFoundException($"Manifest not found: {manifestPath}");
            if (!File.Exists(resultsPath))
                throw new FileNotFoundException($"Results not found: {resultsPath}");

            var manifestText = File.ReadAllText(manifestPath);
            var resultsText = File.ReadAllText(resultsPath);

            var manifest = ParseManifest(manifestText);
            var results = ParseResults(resultsText);

            return CompareFromData(manifest, results);
        }

        /// <summary>
        /// Compare using pre-parsed data (for testing)
        /// </summary>
        public static FullComparisonReport CompareFromData(
            Dictionary<string, object> manifest,
            List<Dictionary<string, object>> results)
        {
            var report = new FullComparisonReport
            {
                RunId = Guid.NewGuid().ToString("N").Substring(0, 8),
                ComparedAt = DateTime.UtcNow
            };

            // Extract default tolerances
            var tolerances = ExtractTolerances(manifest);

            // Get file expectations
            var filesDict = GetDict(manifest, "files");
            if (filesDict == null)
            {
                throw new InvalidOperationException("Manifest missing 'files' section");
            }

            // Compare each result against manifest
            foreach (var result in results)
            {
                var fileName = GetString(result, "FileName");
                if (string.IsNullOrEmpty(fileName))
                    continue;

                var partResult = ComparePartResult(fileName, result, filesDict, tolerances);
                report.Parts.Add(partResult);
            }

            report.CalculateTotals();
            return report;
        }

        private static PartComparisonResult ComparePartResult(
            string fileName,
            Dictionary<string, object> actual,
            Dictionary<string, object> filesDict,
            Dictionary<string, double> tolerances)
        {
            var result = new PartComparisonResult { FileName = fileName };

            // Get expected values for this file
            var fileExpectations = GetDict(filesDict, fileName);
            if (fileExpectations == null)
            {
                result.Fields.Add(new FieldComparison
                {
                    FieldName = "FileInManifest",
                    Status = MatchStatus.MissingExpected,
                    Note = "File not found in manifest"
                });
                result.CalculateCounts();
                return result;
            }

            // Extract expectations with override logic: csharpExpected > vbaBaseline > top-level
            var vbaBaseline = GetDict(fileExpectations, "vbaBaseline") ?? new Dictionary<string, object>();
            var csharpExpected = GetDict(fileExpectations, "csharpExpected") ?? new Dictionary<string, object>();
            var knownDeviations = GetDict(fileExpectations, "knownDeviations") ?? new Dictionary<string, object>();

            // Compare fields
            CompareField(result, "Classification", actual, fileExpectations, csharpExpected, vbaBaseline, knownDeviations, tolerances, "expectedClassification");
            CompareField(result, "Thickness_in", actual, fileExpectations, csharpExpected, vbaBaseline, knownDeviations, tolerances, "expectedThickness_in");
            CompareField(result, "BendCount", actual, fileExpectations, csharpExpected, vbaBaseline, knownDeviations, tolerances, "expectedBendCount");
            CompareField(result, "Material", actual, fileExpectations, csharpExpected, vbaBaseline, knownDeviations, tolerances, "material");
            CompareField(result, "OptiMaterial", actual, fileExpectations, csharpExpected, vbaBaseline, knownDeviations, tolerances, "optiMaterial");
            CompareField(result, "Description", actual, fileExpectations, csharpExpected, vbaBaseline, knownDeviations, tolerances, "description");
            CompareField(result, "MaterialCost", actual, fileExpectations, csharpExpected, vbaBaseline, knownDeviations, tolerances, "materialCost");
            CompareField(result, "FlatArea_sqin", actual, fileExpectations, csharpExpected, vbaBaseline, knownDeviations, tolerances, "flatArea_sqin");
            CompareField(result, "BomQty", actual, fileExpectations, csharpExpected, vbaBaseline, knownDeviations, tolerances, "bomQty");

            // Compare routing if present
            CompareRouting(result, actual, vbaBaseline, csharpExpected, knownDeviations, tolerances);

            result.CalculateCounts();
            return result;
        }

        private static void CompareField(
            PartComparisonResult result,
            string actualFieldName,
            Dictionary<string, object> actual,
            Dictionary<string, object> topLevel,
            Dictionary<string, object> csharpExpected,
            Dictionary<string, object> vbaBaseline,
            Dictionary<string, object> knownDeviations,
            Dictionary<string, double> tolerances,
            string expectedFieldName = null)
        {
            if (expectedFieldName == null)
                expectedFieldName = actualFieldName;

            // Get actual value
            var actualValue = GetValue(actual, actualFieldName);

            // Get expected value (csharp override > vba baseline > top level)
            object expectedValue = null;
            if (csharpExpected.ContainsKey(expectedFieldName))
                expectedValue = csharpExpected[expectedFieldName];
            else if (vbaBaseline.ContainsKey(expectedFieldName))
                expectedValue = vbaBaseline[expectedFieldName];
            else if (topLevel.ContainsKey(expectedFieldName))
                expectedValue = topLevel[expectedFieldName];

            // Check known deviations
            var deviation = GetDict(knownDeviations, expectedFieldName);

            var comparison = new FieldComparison
            {
                FieldName = actualFieldName,
                ExpectedValue = FormatValue(expectedValue),
                ActualValue = FormatValue(actualValue)
            };

            // Handle missing values
            if (actualValue == null || (actualValue is string s && string.IsNullOrWhiteSpace(s)))
            {
                if (deviation != null)
                {
                    var status = GetString(deviation, "status");
                    if (status == "NOT_IMPLEMENTED")
                    {
                        comparison.Status = MatchStatus.NotImplemented;
                        comparison.Note = GetString(deviation, "reason");
                    }
                    else
                    {
                        comparison.Status = MatchStatus.MissingActual;
                        comparison.Note = GetString(deviation, "reason");
                    }
                }
                else
                {
                    comparison.Status = MatchStatus.MissingActual;
                }
                result.Fields.Add(comparison);
                return;
            }

            if (expectedValue == null)
            {
                comparison.Status = MatchStatus.MissingExpected;
                result.Fields.Add(comparison);
                return;
            }

            // Check if this is a known deviation
            if (deviation != null)
            {
                var status = GetString(deviation, "status");
                if (status == "INTENTIONAL_DEVIATION")
                {
                    comparison.Status = MatchStatus.IntentionalDeviation;
                    comparison.Note = GetString(deviation, "reason");
                    result.Fields.Add(comparison);
                    return;
                }
            }

            // Compare values
            if (IsNumeric(actualValue) && IsNumeric(expectedValue))
            {
                CompareNumeric(comparison, actualValue, expectedValue, tolerances, actualFieldName);
            }
            else
            {
                CompareString(comparison, actualValue, expectedValue);
            }

            result.Fields.Add(comparison);
        }

        private static void CompareRouting(
            PartComparisonResult result,
            Dictionary<string, object> actual,
            Dictionary<string, object> vbaBaseline,
            Dictionary<string, object> csharpExpected,
            Dictionary<string, object> knownDeviations,
            Dictionary<string, double> tolerances)
        {
            // Get expected routing (csharp override > vba baseline)
            Dictionary<string, object> expectedRouting = null;
            if (csharpExpected.ContainsKey("routing"))
                expectedRouting = GetDict(csharpExpected, "routing");
            else if (vbaBaseline.ContainsKey("routing"))
                expectedRouting = GetDict(vbaBaseline, "routing");

            if (expectedRouting == null)
                return;

            // Get actual routing — try nested dict first, then build from flat QATestResult fields
            var actualRouting = GetDict(actual, "Routing");
            if (actualRouting == null || actualRouting.Count == 0)
                actualRouting = BuildRoutingFromFlat(actual);

            if (actualRouting == null || actualRouting.Count == 0)
            {
                var deviation = GetDict(knownDeviations, "routing");
                var comparison = new FieldComparison
                {
                    FieldName = "Routing",
                    ExpectedValue = $"{expectedRouting.Count} work centers",
                    ActualValue = "Not populated"
                };

                if (deviation != null && GetString(deviation, "status") == "NOT_IMPLEMENTED")
                {
                    comparison.Status = MatchStatus.NotImplemented;
                    comparison.Note = GetString(deviation, "reason");
                }
                else
                {
                    comparison.Status = MatchStatus.MissingActual;
                }

                result.Fields.Add(comparison);
                return;
            }

            // Compare each work center
            foreach (var kvp in expectedRouting)
            {
                var workCenter = kvp.Key;
                var expectedWc = GetDict(expectedRouting, workCenter);
                var actualWc = GetDict(actualRouting, workCenter);

                if (actualWc == null)
                {
                    result.Fields.Add(new FieldComparison
                    {
                        FieldName = $"Routing.{workCenter}",
                        ExpectedValue = "Present",
                        ActualValue = "Missing",
                        Status = MatchStatus.Fail
                    });
                    continue;
                }

                // Compare op, setup, run
                CompareRoutingField(result, workCenter, "op", expectedWc, actualWc, tolerances);
                CompareRoutingField(result, workCenter, "setup", expectedWc, actualWc, tolerances);
                CompareRoutingField(result, workCenter, "run", expectedWc, actualWc, tolerances);
            }
        }

        private static void CompareRoutingField(
            PartComparisonResult result,
            string workCenter,
            string field,
            Dictionary<string, object> expected,
            Dictionary<string, object> actual,
            Dictionary<string, double> tolerances)
        {
            var expectedValue = GetValue(expected, field);
            var actualValue = GetValue(actual, field);

            var comparison = new FieldComparison
            {
                FieldName = $"Routing.{workCenter}.{field}",
                ExpectedValue = FormatValue(expectedValue),
                ActualValue = FormatValue(actualValue)
            };

            if (actualValue == null)
            {
                comparison.Status = MatchStatus.MissingActual;
            }
            else if (expectedValue == null)
            {
                comparison.Status = MatchStatus.MissingExpected;
            }
            else if (IsNumeric(actualValue) && IsNumeric(expectedValue))
            {
                CompareNumeric(comparison, actualValue, expectedValue, tolerances, "routing");
            }
            else
            {
                CompareString(comparison, actualValue, expectedValue);
            }

            result.Fields.Add(comparison);
        }

        /// <summary>
        /// Build a routing dictionary from flat QATestResult fields (F115_Setup, F140_Run, etc.)
        /// Maps calculator names to ERP work center names used in the manifest.
        /// </summary>
        private static Dictionary<string, object> BuildRoutingFromFlat(Dictionary<string, object> actual)
        {
            var routing = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            // F115 (Laser) → N120 workcenter
            AddRoutingFromFlat(routing, "N120", actual, "F115_Setup", "F115_Run");
            // F140 (Press Brake) → N140 workcenter
            AddRoutingFromFlat(routing, "N140", actual, "F140_Setup", "F140_Run");
            // F210 (Deburr) → N210 workcenter
            AddRoutingFromFlat(routing, "N210", actual, "F210_Setup", "F210_Run");
            // F220 (Tap) → N220 workcenter
            AddRoutingFromFlat(routing, "N220", actual, "F220_Setup", "F220_Run");
            // F325 (Roll Forming) → N325 workcenter
            AddRoutingFromFlat(routing, "N325", actual, "F325_Setup", "F325_Run");
            // F110 (Tube Laser) → F110 workcenter
            AddRoutingFromFlat(routing, "F110", actual, "F110_Setup", "F110_Run");
            // N145 (5-Axis Laser) → N145 workcenter
            AddRoutingFromFlat(routing, "N145", actual, "N145_Setup", "N145_Run");
            // F300 (Saw) → F300 workcenter
            AddRoutingFromFlat(routing, "F300", actual, "F300_Setup", "F300_Run");
            // NPUR (Purchased) → NPUR workcenter
            AddRoutingFromFlat(routing, "NPUR", actual, "NPUR_Setup", "NPUR_Run");
            // CUST (Customer-Supplied) → CUST workcenter
            AddRoutingFromFlat(routing, "CUST", actual, "CUST_Setup", "CUST_Run");

            return routing.Count > 0 ? routing : null;
        }

        private static void AddRoutingFromFlat(
            Dictionary<string, object> routing, string wcKey,
            Dictionary<string, object> actual, string setupField, string runField)
        {
            var setup = GetValue(actual, setupField);
            var run = GetValue(actual, runField);
            if (setup != null || run != null)
            {
                var wc = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                if (setup != null) wc["setup"] = setup;
                if (run != null) wc["run"] = run;
                routing[wcKey] = wc;
            }
        }

        private static void CompareNumeric(
            FieldComparison comparison,
            object actual,
            object expected,
            Dictionary<string, double> tolerances,
            string fieldName)
        {
            var actualNum = ToDouble(actual);
            var expectedNum = ToDouble(expected);

            if (Math.Abs(actualNum - expectedNum) < 1e-10)
            {
                comparison.Status = MatchStatus.Match;
                return;
            }

            // Determine tolerance
            double tolerance = 0.01; // Default 1%
            var lowerField = fieldName.ToLowerInvariant();
            if (lowerField.Contains("thickness") && tolerances.ContainsKey("thickness"))
                tolerance = tolerances["thickness"];
            else if (lowerField.Contains("mass") && tolerances.ContainsKey("mass"))
                tolerance = tolerances["mass"];
            else if (lowerField.Contains("cost") && tolerances.ContainsKey("cost"))
                tolerance = tolerances["cost"];
            else if (lowerField.Contains("routing") && tolerances.ContainsKey("routing"))
                tolerance = tolerances["routing"];
            else if ((lowerField.Contains("area") || lowerField.Contains("dimension")) && tolerances.ContainsKey("dimensions"))
                tolerance = tolerances["dimensions"];

            // Relative difference
            double relativeDiff = Math.Abs(actualNum - expectedNum) / Math.Max(Math.Abs(expectedNum), 1e-10);

            if (relativeDiff <= tolerance)
            {
                comparison.Status = MatchStatus.TolerancePass;
                comparison.Tolerance = $"{tolerance * 100:F1}%";
            }
            else
            {
                comparison.Status = MatchStatus.Fail;
                comparison.Note = $"Difference: {relativeDiff * 100:F2}% (tolerance: {tolerance * 100:F1}%)";
            }
        }

        private static void CompareString(FieldComparison comparison, object actual, object expected)
        {
            var actualStr = (actual?.ToString() ?? string.Empty).Trim();
            var expectedStr = (expected?.ToString() ?? string.Empty).Trim();

            if (string.Equals(actualStr, expectedStr, StringComparison.OrdinalIgnoreCase))
            {
                comparison.Status = MatchStatus.Match;
            }
            else
            {
                comparison.Status = MatchStatus.Fail;
            }
        }

        #region JSON Parsing (Hand-rolled)

        private static Dictionary<string, object> ParseManifest(string json)
        {
            return ParseObject(json);
        }

        private static List<Dictionary<string, object>> ParseResults(string json)
        {
            var root = ParseObject(json);
            var resultsArray = root.ContainsKey("Results") ? root["Results"] as List<object> : null;
            if (resultsArray == null)
                return new List<Dictionary<string, object>>();

            return resultsArray.Cast<Dictionary<string, object>>().ToList();
        }

        private static Dictionary<string, object> ParseObject(string json)
        {
            var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            json = json.Trim();

            if (!json.StartsWith("{") || !json.EndsWith("}"))
                return result;

            json = json.Substring(1, json.Length - 2).Trim();
            if (string.IsNullOrEmpty(json))
                return result;

            var pairs = SplitTopLevel(json, ',');
            foreach (var pair in pairs)
            {
                var colonIndex = FindTopLevelChar(pair, ':');
                if (colonIndex < 0)
                    continue;

                var key = ParseString(pair.Substring(0, colonIndex).Trim());
                var valueStr = pair.Substring(colonIndex + 1).Trim();
                var value = ParseValue(valueStr);

                result[key] = value;
            }

            return result;
        }

        private static List<object> ParseArray(string json)
        {
            var result = new List<object>();
            json = json.Trim();

            if (!json.StartsWith("[") || !json.EndsWith("]"))
                return result;

            json = json.Substring(1, json.Length - 2).Trim();
            if (string.IsNullOrEmpty(json))
                return result;

            var items = SplitTopLevel(json, ',');
            foreach (var item in items)
            {
                result.Add(ParseValue(item.Trim()));
            }

            return result;
        }

        private static object ParseValue(string value)
        {
            value = value.Trim();

            if (value.StartsWith("\""))
                return ParseString(value);
            if (value.StartsWith("{"))
                return ParseObject(value);
            if (value.StartsWith("["))
                return ParseArray(value);
            if (value == "null")
                return null;
            if (value == "true")
                return true;
            if (value == "false")
                return false;

            // Try parse as number
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var num))
                return num;

            return value;
        }

        private static string ParseString(string str)
        {
            str = str.Trim();
            if (str.StartsWith("\"") && str.EndsWith("\""))
                str = str.Substring(1, str.Length - 2);
            return str.Replace("\\\"", "\"").Replace("\\\\", "\\");
        }

        private static List<string> SplitTopLevel(string text, char delimiter)
        {
            var result = new List<string>();
            var current = string.Empty;
            var depth = 0;
            var inString = false;
            var escaped = false;

            for (int i = 0; i < text.Length; i++)
            {
                var c = text[i];

                if (escaped)
                {
                    current += c;
                    escaped = false;
                    continue;
                }

                if (c == '\\' && inString)
                {
                    escaped = true;
                    current += c;
                    continue;
                }

                if (c == '"')
                {
                    inString = !inString;
                    current += c;
                    continue;
                }

                if (inString)
                {
                    current += c;
                    continue;
                }

                if (c == '{' || c == '[')
                {
                    depth++;
                    current += c;
                }
                else if (c == '}' || c == ']')
                {
                    depth--;
                    current += c;
                }
                else if (c == delimiter && depth == 0)
                {
                    if (!string.IsNullOrWhiteSpace(current))
                        result.Add(current);
                    current = string.Empty;
                }
                else
                {
                    current += c;
                }
            }

            if (!string.IsNullOrWhiteSpace(current))
                result.Add(current);

            return result;
        }

        private static int FindTopLevelChar(string text, char target)
        {
            var depth = 0;
            var inString = false;
            var escaped = false;

            for (int i = 0; i < text.Length; i++)
            {
                var c = text[i];

                if (escaped)
                {
                    escaped = false;
                    continue;
                }

                if (c == '\\' && inString)
                {
                    escaped = true;
                    continue;
                }

                if (c == '"')
                {
                    inString = !inString;
                    continue;
                }

                if (inString)
                    continue;

                if (c == '{' || c == '[')
                    depth++;
                else if (c == '}' || c == ']')
                    depth--;
                else if (c == target && depth == 0)
                    return i;
            }

            return -1;
        }

        #endregion

        #region Helper Methods

        private static Dictionary<string, double> ExtractTolerances(Dictionary<string, object> manifest)
        {
            var defaults = new Dictionary<string, double>
            {
                { "thickness", 0.001 },
                { "mass", 0.01 },
                { "dimensions", 0.001 },
                { "cost", 0.05 },
                { "routing", 0.01 }
            };

            var tolerancesDict = GetDict(manifest, "defaultTolerances");
            if (tolerancesDict == null)
                return defaults;

            foreach (var kvp in tolerancesDict)
            {
                if (kvp.Value is double d)
                    defaults[kvp.Key] = d;
                else if (kvp.Value != null && double.TryParse(kvp.Value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
                    defaults[kvp.Key] = parsed;
            }

            return defaults;
        }

        private static Dictionary<string, object> GetDict(Dictionary<string, object> dict, string key)
        {
            if (dict == null || !dict.ContainsKey(key))
                return null;
            return dict[key] as Dictionary<string, object>;
        }

        private static string GetString(Dictionary<string, object> dict, string key)
        {
            if (dict == null || !dict.ContainsKey(key))
                return string.Empty;
            return dict[key]?.ToString() ?? string.Empty;
        }

        private static object GetValue(Dictionary<string, object> dict, string key)
        {
            if (dict == null || !dict.ContainsKey(key))
                return null;
            return dict[key];
        }

        private static bool IsNumeric(object value)
        {
            return value is double || value is int || value is float || value is decimal ||
                   (value is string s && double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out _));
        }

        private static double ToDouble(object value)
        {
            if (value is double d)
                return d;
            if (value is int i)
                return i;
            if (value is float f)
                return f;
            if (value is decimal dec)
                return (double)dec;
            if (value is string s && double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
                return parsed;
            return 0;
        }

        private static string FormatValue(object value)
        {
            if (value == null)
                return string.Empty;
            if (value is double d)
                return d.ToString("F6", CultureInfo.InvariantCulture);
            return value.ToString();
        }

        #endregion
    }
}
