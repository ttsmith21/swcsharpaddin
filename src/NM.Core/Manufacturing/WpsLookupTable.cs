using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace NM.Core.Manufacturing
{
    /// <summary>
    /// Loads and queries the qualified WPS lookup table from a CSV file.
    /// CSV columns: WpsNumber, Process, BaseMetal1, BaseMetal2, ThicknessMinIn, ThicknessMaxIn, JointType, Code, FillerMetal, ShieldingGas, Notes
    /// </summary>
    public sealed class WpsLookupTable
    {
        private readonly List<WpsEntry> _entries = new List<WpsEntry>();

        /// <summary>All loaded WPS entries.</summary>
        public IReadOnlyList<WpsEntry> Entries => _entries;

        /// <summary>True if the table has been loaded with at least one entry.</summary>
        public bool IsLoaded => _entries.Count > 0;

        /// <summary>
        /// Loads WPS entries from a CSV file.
        /// First row is treated as a header and skipped.
        /// Blank lines and lines starting with # are skipped.
        /// </summary>
        public static WpsLookupTable LoadFromCsv(string csvPath)
        {
            var table = new WpsLookupTable();
            if (string.IsNullOrEmpty(csvPath) || !File.Exists(csvPath))
                return table;

            var lines = File.ReadAllLines(csvPath);
            bool headerSkipped = false;

            foreach (var rawLine in lines)
            {
                var line = rawLine.Trim();
                if (string.IsNullOrEmpty(line) || line.StartsWith("#"))
                    continue;

                if (!headerSkipped)
                {
                    headerSkipped = true;
                    continue;
                }

                var entry = ParseLine(line);
                if (entry != null)
                    table._entries.Add(entry);
            }

            return table;
        }

        /// <summary>
        /// Loads WPS entries from an in-memory list (for testing or programmatic use).
        /// </summary>
        public static WpsLookupTable FromEntries(IEnumerable<WpsEntry> entries)
        {
            var table = new WpsLookupTable();
            if (entries != null)
                table._entries.AddRange(entries);
            return table;
        }

        /// <summary>
        /// Finds all WPS entries that match the given joint input.
        /// Material pairs are matched bidirectionally (CS+SS matches SS+CS).
        /// </summary>
        public List<WpsEntry> FindMatches(WpsJointInput input)
        {
            if (input == null) return new List<WpsEntry>();

            string m1 = Normalize(input.BaseMetal1);
            string m2 = Normalize(input.BaseMetal2);
            double t = input.ThicknessIn;
            string jt = Normalize(input.JointType);
            string code = Normalize(input.RequiredCode);

            return _entries.Where(e =>
            {
                // Material pair match (bidirectional)
                string em1 = Normalize(e.BaseMetal1);
                string em2 = Normalize(e.BaseMetal2);
                bool materialMatch = (em1 == m1 && em2 == m2) || (em1 == m2 && em2 == m1);
                if (!materialMatch) return false;

                // Thickness range match
                if (t > 0 && (t < e.ThicknessMinIn || t > e.ThicknessMaxIn))
                    return false;

                // Joint type match (empty input = match all; "Both" in entry matches any)
                if (!string.IsNullOrEmpty(jt) && !string.IsNullOrEmpty(Normalize(e.JointType)))
                {
                    string ejt = Normalize(e.JointType);
                    if (ejt != "both" && ejt != jt)
                        return false;
                }

                // Code match (empty input = match all)
                if (!string.IsNullOrEmpty(code) && !string.IsNullOrEmpty(Normalize(e.Code)))
                {
                    if (Normalize(e.Code) != code)
                        return false;
                }

                return true;
            }).ToList();
        }

        private static string Normalize(string s)
        {
            return string.IsNullOrWhiteSpace(s) ? string.Empty : s.Trim().ToUpperInvariant();
        }

        private static WpsEntry ParseLine(string line)
        {
            // Simple CSV parse (no quoted fields with commas — WPS data doesn't need them)
            var parts = line.Split(',');
            if (parts.Length < 7) return null;

            var entry = new WpsEntry
            {
                WpsNumber = parts[0].Trim(),
                Process = parts.Length > 1 ? parts[1].Trim() : string.Empty,
                BaseMetal1 = parts.Length > 2 ? parts[2].Trim() : string.Empty,
                BaseMetal2 = parts.Length > 3 ? parts[3].Trim() : string.Empty,
                JointType = parts.Length > 6 ? parts[6].Trim() : string.Empty,
                Code = parts.Length > 7 ? parts[7].Trim() : string.Empty,
                FillerMetal = parts.Length > 8 ? parts[8].Trim() : string.Empty,
                ShieldingGas = parts.Length > 9 ? parts[9].Trim() : string.Empty,
                Notes = parts.Length > 10 ? parts[10].Trim() : string.Empty,
            };

            // Parse thickness range
            double tMin, tMax;
            if (parts.Length > 4 && double.TryParse(parts[4].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out tMin))
                entry.ThicknessMinIn = tMin;
            if (parts.Length > 5 && double.TryParse(parts[5].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out tMax))
                entry.ThicknessMaxIn = tMax;

            if (string.IsNullOrEmpty(entry.WpsNumber)) return null;
            return entry;
        }
    }
}
