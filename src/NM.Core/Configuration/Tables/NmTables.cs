using System;
using System.Collections.Generic;

namespace NM.Core.Config.Tables
{
    /// <summary>
    /// Root lookup-tables POCO — deserialized from nm-tables.json.
    /// Replaces both the Excel file and the static fallback classes.
    /// </summary>
    public sealed class NmTables
    {
        public int SchemaVersion { get; set; } = 1;
        public LaserSpeedTable LaserSpeeds { get; set; } = new LaserSpeedTable();
        public TubeCuttingTable TubeCutting { get; set; } = new TubeCuttingTable();
        public List<PipeScheduleEntry> PipeSchedules { get; set; } = new List<PipeScheduleEntry>();
        public GaugeTable Gauges { get; set; } = new GaugeTable();
        public List<MaterialCodeEntry> MaterialCodes { get; set; } = new List<MaterialCodeEntry>();
    }

    // ── Laser speeds ──────────────────────────────────────────────

    public sealed class LaserSpeedTable
    {
        // Group fallbacks (existing — keep for backward compat)
        public List<LaserSpeedEntry> StainlessSteel { get; set; } = new List<LaserSpeedEntry>();
        public List<LaserSpeedEntry> CarbonSteel { get; set; } = new List<LaserSpeedEntry>();
        public List<LaserSpeedEntry> Aluminum { get; set; } = new List<LaserSpeedEntry>();
        public double ThicknessToleranceIn { get; set; } = 0.005;

        // Per-material overrides (checked first, e.g. "304L", "316L", "309", "2205", "CS", "AL")
        public Dictionary<string, List<LaserSpeedEntry>> ByMaterial { get; set; }
            = new Dictionary<string, List<LaserSpeedEntry>>(StringComparer.OrdinalIgnoreCase);
    }

    public sealed class LaserSpeedEntry
    {
        public double ThicknessIn { get; set; }
        public double FeedRateIpm { get; set; }
        public double PierceSeconds { get; set; }
    }

    // ── Tube cutting ──────────────────────────────────────────────

    public sealed class TubeCuttingTable
    {
        public TubeCuttingMaterial CarbonSteel { get; set; } = new TubeCuttingMaterial();
        public TubeCuttingMaterial StainlessSteel { get; set; } = new TubeCuttingMaterial();
        public TubeCuttingMaterial Aluminum { get; set; } = new TubeCuttingMaterial();
    }

    public sealed class TubeCuttingMaterial
    {
        public List<TubeFeedRateEntry> FeedRates { get; set; } = new List<TubeFeedRateEntry>();
        public List<TubePierceTimeEntry> PierceTimes { get; set; } = new List<TubePierceTimeEntry>();
        public double FeedMultiplier { get; set; } = 0.85;
        public double KerfIn { get; set; } = 0.02;
    }

    public sealed class TubeFeedRateEntry
    {
        public double MaxThicknessIn { get; set; }
        public double FeedRateIpm { get; set; }
    }

    public sealed class TubePierceTimeEntry
    {
        public double MaxThicknessIn { get; set; }
        public double PierceSeconds { get; set; }
    }

    // ── Pipe schedules ────────────────────────────────────────────

    public sealed class PipeScheduleEntry
    {
        public double OutsideDiameterIn { get; set; }
        public string NpsText { get; set; }
        public Dictionary<string, double> Schedules { get; set; } = new Dictionary<string, double>();
    }

    // ── Gauge tables ──────────────────────────────────────────────

    public sealed class GaugeTable
    {
        public List<GaugeEntry> StainlessSteel { get; set; } = new List<GaugeEntry>();
        public List<GaugeEntry> ThicknessLabels { get; set; } = new List<GaugeEntry>();
        public double ToleranceIn { get; set; } = 0.008;
    }

    public sealed class GaugeEntry
    {
        public double ThicknessIn { get; set; }
        public string Label { get; set; }
    }

    // ── Material codes ────────────────────────────────────────────

    public sealed class MaterialCodeEntry
    {
        public string SwName { get; set; }
        public string Code { get; set; }
    }
}
