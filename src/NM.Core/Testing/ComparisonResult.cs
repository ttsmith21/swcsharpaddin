using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NM.Core.Testing
{
    /// <summary>
    /// Match status for a field comparison
    /// </summary>
    public enum MatchStatus
    {
        Match,                  // Exact or within tolerance
        TolerancePass,          // Within specified tolerance
        IntentionalDeviation,   // Known C#-vs-VBA difference, documented
        NotImplemented,         // C# doesn't populate this yet
        Fail,                   // Unexpected mismatch
        MissingActual,          // C# didn't produce a value
        MissingExpected         // No baseline value to compare against
    }

    /// <summary>
    /// Single field comparison result
    /// </summary>
    public sealed class FieldComparison
    {
        public string FieldName { get; set; }
        public string ExpectedValue { get; set; }
        public string ActualValue { get; set; }
        public MatchStatus Status { get; set; }
        public string Tolerance { get; set; }  // e.g., "5%" or "0.001"
        public string Note { get; set; }       // e.g., "Intentional: C# uses different tube format"

        public FieldComparison()
        {
            FieldName = string.Empty;
            ExpectedValue = string.Empty;
            ActualValue = string.Empty;
            Status = MatchStatus.Fail;
            Tolerance = string.Empty;
            Note = string.Empty;
        }
    }

    /// <summary>
    /// All field comparisons for one part
    /// </summary>
    public sealed class PartComparisonResult
    {
        public string FileName { get; set; }
        public MatchStatus OverallStatus { get; set; }  // Worst status across all fields
        public List<FieldComparison> Fields { get; set; }
        public int MatchCount { get; set; }
        public int TolerancePassCount { get; set; }
        public int IntentionalDeviationCount { get; set; }
        public int NotImplementedCount { get; set; }
        public int FailCount { get; set; }
        public int MissingCount { get; set; }

        public PartComparisonResult()
        {
            FileName = string.Empty;
            OverallStatus = MatchStatus.Fail;
            Fields = new List<FieldComparison>();
        }

        /// <summary>
        /// Calculate status counts from fields
        /// </summary>
        public void CalculateCounts()
        {
            MatchCount = Fields.Count(f => f.Status == MatchStatus.Match);
            TolerancePassCount = Fields.Count(f => f.Status == MatchStatus.TolerancePass);
            IntentionalDeviationCount = Fields.Count(f => f.Status == MatchStatus.IntentionalDeviation);
            NotImplementedCount = Fields.Count(f => f.Status == MatchStatus.NotImplemented);
            FailCount = Fields.Count(f => f.Status == MatchStatus.Fail);
            MissingCount = Fields.Count(f => f.Status == MatchStatus.MissingActual || f.Status == MatchStatus.MissingExpected);

            // Determine worst status
            if (FailCount > 0)
                OverallStatus = MatchStatus.Fail;
            else if (MissingCount > 0)
                OverallStatus = MatchStatus.MissingActual;
            else if (NotImplementedCount > 0)
                OverallStatus = MatchStatus.NotImplemented;
            else if (IntentionalDeviationCount > 0)
                OverallStatus = MatchStatus.IntentionalDeviation;
            else if (TolerancePassCount > 0)
                OverallStatus = MatchStatus.TolerancePass;
            else
                OverallStatus = MatchStatus.Match;
        }
    }

    /// <summary>
    /// Full comparison report for all parts
    /// </summary>
    public sealed class FullComparisonReport
    {
        public string RunId { get; set; }
        public DateTime ComparedAt { get; set; }
        public List<PartComparisonResult> Parts { get; set; }
        public int TotalFields { get; set; }
        public int TotalMatch { get; set; }
        public int TotalTolerancePass { get; set; }
        public int TotalDeviation { get; set; }
        public int TotalNotImplemented { get; set; }
        public int TotalFail { get; set; }

        public FullComparisonReport()
        {
            RunId = string.Empty;
            ComparedAt = DateTime.UtcNow;
            Parts = new List<PartComparisonResult>();
        }

        /// <summary>
        /// Calculate aggregate statistics
        /// </summary>
        public void CalculateTotals()
        {
            TotalFields = Parts.Sum(p => p.Fields.Count);
            TotalMatch = Parts.Sum(p => p.MatchCount);
            TotalTolerancePass = Parts.Sum(p => p.TolerancePassCount);
            TotalDeviation = Parts.Sum(p => p.IntentionalDeviationCount);
            TotalNotImplemented = Parts.Sum(p => p.NotImplementedCount);
            TotalFail = Parts.Sum(p => p.FailCount);
        }

        /// <summary>
        /// Generates a human-readable summary string
        /// </summary>
        public string GenerateSummary()
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== GOLD STANDARD COMPARISON SUMMARY ===");
            sb.AppendLine($"Run ID: {RunId}");
            sb.AppendLine($"Compared At: {ComparedAt:yyyy-MM-dd HH:mm:ss} UTC");
            sb.AppendLine();
            sb.AppendLine($"Parts Compared: {Parts.Count}");
            sb.AppendLine($"Total Fields: {TotalFields}");
            sb.AppendLine();
            sb.AppendLine("Field Status Breakdown:");
            sb.AppendLine($"  ✓ Match:            {TotalMatch,5} ({Percentage(TotalMatch, TotalFields),5:F1}%)");
            sb.AppendLine($"  ≈ Tolerance Pass:   {TotalTolerancePass,5} ({Percentage(TotalTolerancePass, TotalFields),5:F1}%)");
            sb.AppendLine($"  ⚠ Known Deviation:  {TotalDeviation,5} ({Percentage(TotalDeviation, TotalFields),5:F1}%)");
            sb.AppendLine($"  ○ Not Implemented:  {TotalNotImplemented,5} ({Percentage(TotalNotImplemented, TotalFields),5:F1}%)");
            sb.AppendLine($"  ✗ FAIL:             {TotalFail,5} ({Percentage(TotalFail, TotalFields),5:F1}%)");
            sb.AppendLine();

            int passCount = Parts.Count(p => p.OverallStatus == MatchStatus.Match || p.OverallStatus == MatchStatus.TolerancePass);
            int deviationCount = Parts.Count(p => p.OverallStatus == MatchStatus.IntentionalDeviation);
            int notImplCount = Parts.Count(p => p.OverallStatus == MatchStatus.NotImplemented);
            int failCount = Parts.Count(p => p.OverallStatus == MatchStatus.Fail || p.OverallStatus == MatchStatus.MissingActual);

            sb.AppendLine("Part-Level Status:");
            sb.AppendLine($"  ✓ Pass:            {passCount,3}/{Parts.Count}");
            sb.AppendLine($"  ⚠ Known Deviation: {deviationCount,3}/{Parts.Count}");
            sb.AppendLine($"  ○ Not Implemented: {notImplCount,3}/{Parts.Count}");
            sb.AppendLine($"  ✗ FAIL:            {failCount,3}/{Parts.Count}");

            return sb.ToString();
        }

        /// <summary>
        /// Generates detailed per-part report
        /// </summary>
        public string GenerateDetailedReport()
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== DETAILED COMPARISON REPORT ===");
            sb.AppendLine();

            foreach (var part in Parts.OrderBy(p => p.FileName))
            {
                sb.AppendLine($"File: {part.FileName}");
                sb.AppendLine($"Overall Status: {part.OverallStatus}");
                sb.AppendLine($"Fields: {part.Fields.Count} total ({part.MatchCount} match, {part.TolerancePassCount} tolerance, {part.IntentionalDeviationCount} deviation, {part.NotImplementedCount} not impl, {part.FailCount} fail)");
                sb.AppendLine();

                // Group by status for readability
                var failedFields = part.Fields.Where(f => f.Status == MatchStatus.Fail).ToList();
                var missingFields = part.Fields.Where(f => f.Status == MatchStatus.MissingActual || f.Status == MatchStatus.MissingExpected).ToList();
                var notImplFields = part.Fields.Where(f => f.Status == MatchStatus.NotImplemented).ToList();
                var deviationFields = part.Fields.Where(f => f.Status == MatchStatus.IntentionalDeviation).ToList();
                var toleranceFields = part.Fields.Where(f => f.Status == MatchStatus.TolerancePass).ToList();

                if (failedFields.Any())
                {
                    sb.AppendLine("  FAILURES:");
                    foreach (var field in failedFields)
                    {
                        sb.AppendLine($"    ✗ {field.FieldName}");
                        sb.AppendLine($"      Expected: {field.ExpectedValue}");
                        sb.AppendLine($"      Actual:   {field.ActualValue}");
                        if (!string.IsNullOrEmpty(field.Note))
                            sb.AppendLine($"      Note:     {field.Note}");
                    }
                    sb.AppendLine();
                }

                if (missingFields.Any())
                {
                    sb.AppendLine("  MISSING:");
                    foreach (var field in missingFields)
                    {
                        sb.AppendLine($"    ? {field.FieldName}: {field.Status}");
                        if (!string.IsNullOrEmpty(field.Note))
                            sb.AppendLine($"      Note: {field.Note}");
                    }
                    sb.AppendLine();
                }

                if (notImplFields.Any())
                {
                    sb.AppendLine("  NOT IMPLEMENTED:");
                    foreach (var field in notImplFields)
                    {
                        sb.AppendLine($"    ○ {field.FieldName}");
                        if (!string.IsNullOrEmpty(field.Note))
                            sb.AppendLine($"      Note: {field.Note}");
                    }
                    sb.AppendLine();
                }

                if (deviationFields.Any())
                {
                    sb.AppendLine("  KNOWN DEVIATIONS:");
                    foreach (var field in deviationFields)
                    {
                        sb.AppendLine($"    ⚠ {field.FieldName}");
                        sb.AppendLine($"      Expected: {field.ExpectedValue}");
                        sb.AppendLine($"      Actual:   {field.ActualValue}");
                        sb.AppendLine($"      Note:     {field.Note}");
                    }
                    sb.AppendLine();
                }

                if (toleranceFields.Any())
                {
                    sb.AppendLine("  TOLERANCE PASS:");
                    foreach (var field in toleranceFields)
                    {
                        sb.AppendLine($"    ≈ {field.FieldName} (tolerance: {field.Tolerance})");
                        sb.AppendLine($"      Expected: {field.ExpectedValue}");
                        sb.AppendLine($"      Actual:   {field.ActualValue}");
                    }
                    sb.AppendLine();
                }

                sb.AppendLine(new string('-', 80));
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private double Percentage(int value, int total)
        {
            if (total == 0) return 0;
            return (value * 100.0) / total;
        }
    }
}
