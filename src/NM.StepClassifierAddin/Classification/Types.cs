using System;
using SolidWorks.Interop.sldworks;

namespace NM.StepClassifierAddin.Classification
{
    /// <summary>Piles returned by the classifier.</summary>
    public enum PartPile
    {
        Stick,
        SheetMetal,
        Other
    }

    /// <summary>Principal oriented spans (L1 ? L2 ? L3) in meters.</summary>
    public struct OrientedSpans
    {
        public double L1, L2, L3;
    }

    /// <summary>Axis alignment and area shares for a principal axis.</summary>
    public sealed class AxisStats
    {
        public double[] Axis = new double[3];
        public double SideArea; // area of sides aligned with axis
        public double CapArea;  // area of caps aligned with axis
    }

    /// <summary>Aggregated metrics from the prepass scan.</summary>
    public sealed class PrepassMetrics
    {
        public double ATotal;
        public double ADevelopable;
        public AxisStats[] Axes = new AxisStats[] { new AxisStats(), new AxisStats(), new AxisStats() };
        public OrientedSpans Spans;
        public IBody2 Body; // for trace
    }

    /// <summary>Classifier thresholds and constants.</summary>
    public static class Thresholds
    {
        public const double ALIGN_TOL = 0.05;            // ?3°
        public const double NORMAL_PERP_TOL = 0.15;      // ?8–9°
        public const double CAP_ALIGN_TOL = 0.05;
        public const double DEV_SURF_MIN_SHARE = 0.85;
        public const double STICK_SIDE_MIN_SHARE = 0.80;
        public const double STICK_CAP_MIN_SHARE = 0.05;
        public const double SHEET_THICK_COVERAGE = 0.70;
        public const double SHEET_THIN_RATIO_L2 = 0.20;
        public const double SECTION_AREA_TOL = 0.05;     // 5%
    }
}