using System;
using System.Collections.Generic;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Lookup tables for ISO 13920 (welding) and ISO 2768 (general machining/laser)
    /// tolerance standards. Used by FabricationToleranceClassifier to evaluate
    /// whether a drawing's tolerances are tighter than the shop standard.
    ///
    /// ISO 13920:2023 — General tolerances for welded constructions
    ///   Linear classes: A (fine), B (medium), C (coarse), D (very coarse)
    ///   Flatness/Straightness classes: E (fine), F (medium), G (coarse), H (very coarse)
    ///
    /// ISO 2768-1:1989 — General tolerances for linear and angular dimensions
    ///   Classes: f (fine), m (medium), c (coarse), v (very coarse)
    /// </summary>
    public static class IsoToleranceStandard
    {
        // =====================================================================
        // ISO 13920 — Welding tolerances (mm)
        // =====================================================================

        /// <summary>ISO 13920 linear dimension tolerance classes.</summary>
        public enum Iso13920Linear { A, B, C, D }

        /// <summary>ISO 13920 flatness/straightness/parallelism tolerance classes.</summary>
        public enum Iso13920Geometric { E, F, G, H }

        // Table 1: Linear dimension tolerances (± mm)
        // Rows: nominal size breakpoints (upper bound in mm)
        // Columns: A, B, C, D
        private static readonly (double MaxMm, double A, double B, double C, double D)[] Iso13920LinearTable =
        {
            (30,     1,  2,  3,  4),
            (120,    1,  2,  4,  7),
            (400,    1,  3,  6,  9),
            (1000,   2,  4,  8,  12),
            (2000,   3,  6,  11, 16),
            (4000,   4,  8,  14, 21),
            (8000,   5,  10, 18, 27),
            (12000,  6,  12, 21, 32),
            (16000,  7,  14, 24, 36),
            (20000,  8,  16, 27, 40),
        };

        // Table 3: Flatness/Straightness/Parallelism tolerances (mm, NOT ±)
        private static readonly (double MaxMm, double E, double F, double G, double H)[] Iso13920GeometricTable =
        {
            (120,    0.5, 1,   1.5, 2.5),
            (400,    1,   1.5, 3,   5),
            (1000,   1.5, 3,   5.5, 9),
            (2000,   2,   4.5, 9,   14),
            (4000,   3,   6,   11,  18),
            (8000,   4,   8,   16,  26),
            (12000,  5,   10,  20,  32),
            (16000,  6,   12,  22,  36),
            (20000,  7,   14,  25,  40),
        };

        /// <summary>
        /// Gets the ISO 13920 linear tolerance (± mm) for a given nominal size and class.
        /// </summary>
        public static double GetIso13920Linear(double nominalMm, Iso13920Linear cls)
        {
            foreach (var row in Iso13920LinearTable)
            {
                if (nominalMm <= row.MaxMm)
                {
                    switch (cls)
                    {
                        case Iso13920Linear.A: return row.A;
                        case Iso13920Linear.B: return row.B;
                        case Iso13920Linear.C: return row.C;
                        case Iso13920Linear.D: return row.D;
                    }
                }
            }
            // Beyond table range, use last row
            var last = Iso13920LinearTable[Iso13920LinearTable.Length - 1];
            switch (cls)
            {
                case Iso13920Linear.A: return last.A;
                case Iso13920Linear.B: return last.B;
                case Iso13920Linear.C: return last.C;
                default: return last.D;
            }
        }

        /// <summary>
        /// Gets the ISO 13920 flatness/straightness tolerance (mm) for a given length and class.
        /// </summary>
        public static double GetIso13920Geometric(double lengthMm, Iso13920Geometric cls)
        {
            foreach (var row in Iso13920GeometricTable)
            {
                if (lengthMm <= row.MaxMm)
                {
                    switch (cls)
                    {
                        case Iso13920Geometric.E: return row.E;
                        case Iso13920Geometric.F: return row.F;
                        case Iso13920Geometric.G: return row.G;
                        case Iso13920Geometric.H: return row.H;
                    }
                }
            }
            var last = Iso13920GeometricTable[Iso13920GeometricTable.Length - 1];
            switch (cls)
            {
                case Iso13920Geometric.E: return last.E;
                case Iso13920Geometric.F: return last.F;
                case Iso13920Geometric.G: return last.G;
                default: return last.H;
            }
        }

        /// <summary>Returns true if classA is tighter (finer) than classB.</summary>
        public static bool IsTighterLinear(Iso13920Linear a, Iso13920Linear b) => a < b;

        /// <summary>Returns true if classA is tighter (finer) than classB.</summary>
        public static bool IsTighterGeometric(Iso13920Geometric a, Iso13920Geometric b) => a < b;

        // =====================================================================
        // ISO 2768-1 — General tolerances for linear dimensions (mm)
        // =====================================================================

        /// <summary>ISO 2768-1 linear tolerance classes.</summary>
        public enum Iso2768Linear { Fine, Medium, Coarse, VeryCoarse }

        // Table 1: Permissible deviations (± mm)
        private static readonly (double MaxMm, double F, double M, double C, double V)[] Iso2768LinearTable =
        {
            (3,      0.05, 0.1,  0.2,  double.NaN), // v not defined for <3
            (6,      0.05, 0.1,  0.3,  0.5),
            (30,     0.1,  0.2,  0.5,  1.0),
            (120,    0.15, 0.3,  0.8,  1.5),
            (400,    0.2,  0.5,  1.2,  2.5),
            (1000,   0.3,  0.8,  2.0,  4.0),
            (2000,   0.5,  1.2,  3.0,  6.0),
            (4000,   double.NaN, 2.0, 4.0, 8.0), // f not defined for >2000
        };

        /// <summary>
        /// Gets the ISO 2768-1 linear tolerance (± mm) for a given nominal size and class.
        /// Returns NaN if the class is not defined for this size range.
        /// </summary>
        public static double GetIso2768Linear(double nominalMm, Iso2768Linear cls)
        {
            foreach (var row in Iso2768LinearTable)
            {
                if (nominalMm <= row.MaxMm)
                {
                    switch (cls)
                    {
                        case Iso2768Linear.Fine: return row.F;
                        case Iso2768Linear.Medium: return row.M;
                        case Iso2768Linear.Coarse: return row.C;
                        case Iso2768Linear.VeryCoarse: return row.V;
                    }
                }
            }
            var last = Iso2768LinearTable[Iso2768LinearTable.Length - 1];
            switch (cls)
            {
                case Iso2768Linear.Fine: return last.F;
                case Iso2768Linear.Medium: return last.M;
                case Iso2768Linear.Coarse: return last.C;
                default: return last.V;
            }
        }

        /// <summary>Returns true if classA is tighter (finer) than classB.</summary>
        public static bool IsTighterLinear(Iso2768Linear a, Iso2768Linear b) => a < b;

        // =====================================================================
        // Unit conversion helpers
        // =====================================================================

        private const double MmPerInch = 25.4;

        /// <summary>Converts inches to millimeters.</summary>
        public static double InchesToMm(double inches) => inches * MmPerInch;

        /// <summary>Converts millimeters to inches.</summary>
        public static double MmToInches(double mm) => mm / MmPerInch;

        /// <summary>
        /// Determines which ISO 13920 linear class a given tolerance band falls into.
        /// Returns the tightest class that the tolerance satisfies (i.e., is looser than or equal to).
        /// </summary>
        public static Iso13920Linear? ClassifyLinearTolerance13920(double nominalMm, double toleranceBandMm)
        {
            double halfBand = toleranceBandMm / 2.0; // ISO tables use ± values
            double tolA = GetIso13920Linear(nominalMm, Iso13920Linear.A);
            double tolB = GetIso13920Linear(nominalMm, Iso13920Linear.B);
            double tolC = GetIso13920Linear(nominalMm, Iso13920Linear.C);
            double tolD = GetIso13920Linear(nominalMm, Iso13920Linear.D);

            if (halfBand <= tolA) return Iso13920Linear.A;
            if (halfBand <= tolB) return Iso13920Linear.B;
            if (halfBand <= tolC) return Iso13920Linear.C;
            if (halfBand <= tolD) return Iso13920Linear.D;
            return null; // Looser than class D — no class needed
        }

        /// <summary>
        /// Determines which ISO 13920 geometric class a given flatness/straightness tolerance falls into.
        /// </summary>
        public static Iso13920Geometric? ClassifyGeometricTolerance13920(double lengthMm, double toleranceMm)
        {
            double tolE = GetIso13920Geometric(lengthMm, Iso13920Geometric.E);
            double tolF = GetIso13920Geometric(lengthMm, Iso13920Geometric.F);
            double tolG = GetIso13920Geometric(lengthMm, Iso13920Geometric.G);
            double tolH = GetIso13920Geometric(lengthMm, Iso13920Geometric.H);

            if (toleranceMm <= tolE) return Iso13920Geometric.E;
            if (toleranceMm <= tolF) return Iso13920Geometric.F;
            if (toleranceMm <= tolG) return Iso13920Geometric.G;
            if (toleranceMm <= tolH) return Iso13920Geometric.H;
            return null;
        }
    }
}
