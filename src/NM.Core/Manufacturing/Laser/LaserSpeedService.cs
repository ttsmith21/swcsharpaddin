using System;
using System.Collections.Generic;
using System.Globalization;

namespace NM.Core.Manufacturing.Laser
{
    // Represents one row from the Excel laser sheet
    public sealed class LaserSpeedRow
    {
        public double ThicknessIn { get; set; } // inches
        // Speeds and pierce times are per material column mapping
        public Dictionary<int, double> ColumnValue { get; } = new Dictionary<int, double>(); // column -> value
    }

    // Result of speed lookup
    public struct LaserSpeed
    {
        public double FeedRateIpm;   // inches per minute
        public double PierceSeconds; // seconds per pierce
        public bool HasValue => FeedRateIpm > 0 && PierceSeconds >= 0;
    }

    public interface ILaserSpeedProvider
    {
        // Returns feed and pierce for given thickness/material; must implement fuzzy matching rules
        LaserSpeed GetSpeed(double thicknessIn, string materialCode);
    }

    // In-memory table provider (populate from Excel elsewhere)
    public sealed class LaserSpeedTableProvider : ILaserSpeedProvider
    {
        private readonly List<LaserSpeedRow> _rows;
        public LaserSpeedTableProvider(IEnumerable<LaserSpeedRow> rows)
        {
            _rows = new List<LaserSpeedRow>(rows ?? Array.Empty<LaserSpeedRow>());
            _rows.Sort((a, b) => a.ThicknessIn.CompareTo(b.ThicknessIn));
        }

        private const double THICKNESS_TOLERANCE = 0.005; // inches

        // Excel column mapping per material (19..22)
        private static int GetSpeedColumn(string material)
        {
            var m = (material ?? string.Empty).ToUpperInvariant();
            if (m.Contains("304") || m.Contains("316")) return 19; // S
            if (m.Contains("309") || m.Contains("2205")) return 20; // T
            if (m.Contains("A36") || m.Contains("ALNZD") || m.Contains("CS")) return 21; // U
            if (m.Contains("6061") || m.Contains("5052") || m.Contains("AL")) return 22; // V
            return 19;
        }

        public LaserSpeed GetSpeed(double thicknessIn, string materialCode)
        {
            int col = GetSpeedColumn(materialCode);
            if (_rows.Count == 0) return default;

            // 1) exact match within tolerance
            foreach (var r in _rows)
            {
                if (Math.Abs(r.ThicknessIn - thicknessIn) <= THICKNESS_TOLERANCE)
                {
                    return BuildSpeed(r, col);
                }
            }

            // 2) nearest greater thickness
            foreach (var r in _rows)
            {
                if (r.ThicknessIn > thicknessIn)
                {
                    return BuildSpeed(r, col);
                }
            }

            // 3) maximum thickness row
            return BuildSpeed(_rows[_rows.Count - 1], col);
        }

        private static LaserSpeed BuildSpeed(LaserSpeedRow row, int column)
        {
            // Expect two adjacent columns for feed and pierce? The VBA maps 19-22 to feeds; pierce often in a neighboring column.
            // Here we assume: speed column is feed IPM; pierce is provided in column+?; until Excel schema is wired, return Feed only.
            double feed = row.ColumnValue.TryGetValue(column, out var v) ? v : 0;
            double pierce = 0; // Will be populated when Excel loader includes pierce column
            return new LaserSpeed { FeedRateIpm = feed, PierceSeconds = pierce };
        }
    }
}
