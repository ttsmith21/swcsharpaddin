using System;
using System.Globalization;
using NM.Core.Manufacturing.Laser;
using NM.SwAddin.Data;

namespace NM.SwAddin.Manufacturing.Laser
{
    // ILaserSpeedProvider implementation backed by ExcelDataLoader tables
    public sealed class LaserSpeedExcelProvider : ILaserSpeedProvider
    {
        private readonly ExcelDataLoader _loader;
        public LaserSpeedExcelProvider(ExcelDataLoader loader)
        {
            _loader = loader ?? throw new ArgumentNullException(nameof(loader));
        }

        // Column and row indices (1-based per Excel)
        private const int THICKNESS_COLUMN = 3; // C
        private const int HEADER_ROW = 1;
        private const int DATA_START_ROW = 2;
        private const double TOL = 0.005; // inches

        public LaserSpeed GetSpeed(double thicknessIn, string materialCode)
        {
            var sheet = ResolveWorksheet(materialCode);
            if (sheet == null) return default;

            int speedCol = GetSpeedColumn(materialCode);
            int pierceCol = GetPierceColumn(materialCode);

            var (lbR, ubR, lbC, ubC) = GetBounds(sheet);

            // 1) exact match within tolerance
            int bestRow = -1;
            for (int r = Math.Max(lbR, DATA_START_ROW); r <= ubR; r++)
            {
                double t = ToDouble(sheet[r, THICKNESS_COLUMN]);
                if (Math.Abs(t - thicknessIn) <= TOL)
                {
                    bestRow = r; break;
                }
            }

            // 2) nearest greater
            if (bestRow < 0)
            {
                double minDiff = double.MaxValue;
                for (int r = Math.Max(lbR, DATA_START_ROW); r <= ubR; r++)
                {
                    double t = ToDouble(sheet[r, THICKNESS_COLUMN]);
                    if (t > thicknessIn)
                    {
                        double diff = t - thicknessIn;
                        if (diff < minDiff)
                        {
                            minDiff = diff; bestRow = r;
                        }
                    }
                }
            }

            // 3) max thickness row
            if (bestRow < 0) bestRow = ubR;

            double feed = ToDouble(sheet[bestRow, speedCol]);
            double pierce = ToDouble(sheet[bestRow, pierceCol]);
            return new LaserSpeed { FeedRateIpm = feed, PierceSeconds = pierce };
        }

        private object[,] ResolveWorksheet(string materialCode)
        {
            // Try exact tabs first
            var m = (materialCode ?? string.Empty).ToUpperInvariant();
            string tab = null;
            if (m.Contains("304")) tab = "304L";
            else if (m.Contains("316")) tab = "316L";
            else if (m.Contains("309")) tab = "309";
            else if (m.Contains("2205")) tab = "2205";
            else if (m.Contains("A36") || m.Contains("ALNZD") || m.Contains("CS")) tab = "CS";
            else if (m.Contains("6061") || m.Contains("5052") || m.Contains("AL")) tab = "AL";

            if (!string.IsNullOrEmpty(tab))
            {
                var ws = _loader.GetWorksheet(tab);
                if (ws != null) return ws;
            }

            // Fallback to generic grouped sheets
            if (m.Contains("A36") || m.Contains("ALNZD") || m.Contains("CS")) return _loader.CarbonLaserSpeedsNFeeds;
            if (m.Contains("6061") || m.Contains("5052") || m.Contains("AL")) return _loader.AluminumLaserSpeedsNFeeds;
            return _loader.SSLaserSpeedsNFeeds; // default stainless
        }

        private static (int lbR, int ubR, int lbC, int ubC) GetBounds(object[,] arr)
        {
            return (arr.GetLowerBound(0), arr.GetUpperBound(0), arr.GetLowerBound(1), arr.GetUpperBound(1));
        }

        private static int GetSpeedColumn(string material)
        {
            var m = (material ?? string.Empty).ToUpperInvariant();
            if (m.Contains("304") || m.Contains("316")) return 19; // S
            if (m.Contains("309") || m.Contains("2205")) return 20; // T
            if (m.Contains("A36") || m.Contains("ALNZD") || m.Contains("CS")) return 21; // U
            if (m.Contains("6061") || m.Contains("5052") || m.Contains("AL")) return 22; // V
            return 19;
        }

        private static int GetPierceColumn(string material)
        {
            var m = (material ?? string.Empty).ToUpperInvariant();
            if (m.Contains("304") || m.Contains("316")) return 23; // W
            if (m.Contains("309") || m.Contains("2205")) return 24; // X
            if (m.Contains("A36") || m.Contains("ALNZD") || m.Contains("CS")) return 25; // Y
            if (m.Contains("6061") || m.Contains("5052") || m.Contains("AL")) return 26; // Z
            return 23;
        }

        private static double ToDouble(object cell)
        {
            if (cell == null) return 0;
            if (cell is double d) return d;
            if (double.TryParse(cell.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) return v;
            if (double.TryParse(cell.ToString(), out v)) return v;
            return 0;
        }
    }
}
