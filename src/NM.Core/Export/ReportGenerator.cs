using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace NM.Core.Export
{
    /// <summary>
    /// Generates summary reports for processed parts and assemblies.
    /// Ported from VBA SP.bas Report() and ReportPart() functions.
    /// Outputs CSV format for easy Excel import.
    /// </summary>
    public sealed class ReportGenerator
    {
        /// <summary>
        /// Data for a single part in the report.
        /// </summary>
        public sealed class PartReportData
        {
            public string PartName { get; set; }
            public string FilePath { get; set; }
            public double WeightLb { get; set; }
            public string OptiMaterial { get; set; }
            public double RawWeight { get; set; }
            public string OP20 { get; set; }  // Primary work center
            public double OP20_Setup { get; set; }
            public double OP20_Run { get; set; }
            public bool HasPressbrake { get; set; }
            public double F140_Setup { get; set; }
            public double F140_Run { get; set; }
            public double F210_Setup { get; set; }
            public double F210_Run { get; set; }
            public double F220_Setup { get; set; }
            public double F220_Run { get; set; }
            public double F325_Setup { get; set; }
            public double F325_Run { get; set; }
            public double Other_Setup { get; set; }
            public double Other_Run { get; set; }
            public int Quantity { get; set; } = 1;
            public string PartType { get; set; }  // Manufactured, Purchased, Customer-Supplied
            public string Notes { get; set; }
        }

        /// <summary>
        /// Assembly report containing multiple parts with quantities.
        /// </summary>
        public sealed class AssemblyReportData
        {
            public string AssemblyName { get; set; }
            public string FilePath { get; set; }
            public DateTime GeneratedAt { get; set; } = DateTime.Now;
            public List<PartReportData> Parts { get; } = new List<PartReportData>();
        }

        /// <summary>
        /// Report for standalone parts in a folder.
        /// </summary>
        public sealed class FolderReportData
        {
            public string FolderPath { get; set; }
            public DateTime GeneratedAt { get; set; } = DateTime.Now;
            public List<PartReportData> Parts { get; } = new List<PartReportData>();
        }

        /// <summary>
        /// Generates a CSV report for an assembly's BOM.
        /// Equivalent to VBA Report() function.
        /// </summary>
        public void GenerateAssemblyReport(AssemblyReportData data, string outputPath)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (string.IsNullOrEmpty(outputPath)) throw new ArgumentNullException(nameof(outputPath));

            var sb = new StringBuilder();

            // Header
            sb.AppendLine($"Assembly Report: {data.AssemblyName}");
            sb.AppendLine($"Generated: {data.GeneratedAt:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"Source: {data.FilePath}");
            sb.AppendLine();

            // Column headers
            sb.AppendLine(GetCsvHeader());

            // Data rows
            foreach (var part in data.Parts)
            {
                sb.AppendLine(GetCsvRow(part));
            }

            // Summary section
            sb.AppendLine();
            sb.AppendLine(GenerateSummary(data.Parts));

            File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// Generates a CSV report for parts in a folder.
        /// Equivalent to VBA ReportPart() function.
        /// </summary>
        public void GenerateFolderReport(FolderReportData data, string outputPath)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (string.IsNullOrEmpty(outputPath)) throw new ArgumentNullException(nameof(outputPath));

            var sb = new StringBuilder();

            // Header
            sb.AppendLine($"Parts Report");
            sb.AppendLine($"Generated: {data.GeneratedAt:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"Folder: {data.FolderPath}");
            sb.AppendLine();

            // Column headers
            sb.AppendLine(GetCsvHeader());

            // Data rows
            foreach (var part in data.Parts)
            {
                sb.AppendLine(GetCsvRow(part));
            }

            // Summary section
            sb.AppendLine();
            sb.AppendLine(GenerateSummary(data.Parts));

            File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);
        }

        private static string GetCsvHeader()
        {
            return string.Join(",",
                "Part Name",
                "Weight (lb)",
                "OptiMaterial",
                "Raw Weight/Length",
                "OP20",
                "OP20 Setup (hr)",
                "OP20 Run (hr)",
                "Pressbrake",
                "F140 Setup (hr)",
                "F140 Run (hr)",
                "Other Setup (hr)",
                "Other Run (hr)",
                "Qty",
                "Part Type",
                "Notes"
            );
        }

        private static string GetCsvRow(PartReportData part)
        {
            return string.Join(",",
                CsvEscape(part.PartName),
                part.WeightLb.ToString("F3", CultureInfo.InvariantCulture),
                CsvEscape(part.OptiMaterial),
                part.RawWeight.ToString("F3", CultureInfo.InvariantCulture),
                CsvEscape(part.OP20),
                part.OP20_Setup.ToString("F4", CultureInfo.InvariantCulture),
                part.OP20_Run.ToString("F4", CultureInfo.InvariantCulture),
                part.HasPressbrake ? "Y" : "N",
                part.F140_Setup.ToString("F4", CultureInfo.InvariantCulture),
                part.F140_Run.ToString("F4", CultureInfo.InvariantCulture),
                (part.F210_Setup + part.F220_Setup + part.F325_Setup + part.Other_Setup).ToString("F4", CultureInfo.InvariantCulture),
                (part.F210_Run + part.F220_Run + part.F325_Run + part.Other_Run).ToString("F4", CultureInfo.InvariantCulture),
                part.Quantity.ToString(),
                CsvEscape(part.PartType),
                CsvEscape(part.Notes)
            );
        }

        private static string CsvEscape(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }
            return value;
        }

        private static string GenerateSummary(List<PartReportData> parts)
        {
            if (parts == null || parts.Count == 0)
                return "No parts in report.";

            var sb = new StringBuilder();
            sb.AppendLine("=== SUMMARY ===");
            sb.AppendLine($"Total Parts: {parts.Count}");

            double totalWeight = 0;
            double totalSetup = 0;
            double totalRun = 0;
            int manufacturedCount = 0;
            int purchasedCount = 0;

            foreach (var part in parts)
            {
                totalWeight += part.WeightLb * part.Quantity;
                totalSetup += part.OP20_Setup + part.F140_Setup + part.F210_Setup + part.F220_Setup + part.F325_Setup + part.Other_Setup;
                totalRun += (part.OP20_Run + part.F140_Run + part.F210_Run + part.F220_Run + part.F325_Run + part.Other_Run) * part.Quantity;

                if (part.PartType == "Purchased" || part.PartType == "Customer-Supplied")
                    purchasedCount++;
                else
                    manufacturedCount++;
            }

            sb.AppendLine($"Manufactured Parts: {manufacturedCount}");
            sb.AppendLine($"Purchased/Supplied Parts: {purchasedCount}");
            sb.AppendLine($"Total Weight: {totalWeight:F2} lb");
            sb.AppendLine($"Total Setup Time: {totalSetup:F2} hr");
            sb.AppendLine($"Total Run Time: {totalRun:F2} hr");
            sb.AppendLine($"Total Production Time: {totalSetup + totalRun:F2} hr");

            return sb.ToString();
        }

        /// <summary>
        /// Creates a PartReportData from custom properties dictionary.
        /// </summary>
        public static PartReportData FromCustomProperties(string partName, Dictionary<string, string> props)
        {
            var data = new PartReportData { PartName = partName };

            if (props == null) return data;

            data.WeightLb = GetDouble(props, "Weight");
            data.OptiMaterial = GetString(props, "OptiMaterial");
            data.RawWeight = GetDouble(props, "RawWeight");
            if (data.RawWeight == 0)
                data.RawWeight = GetDouble(props, "F300_Length");

            data.OP20 = GetString(props, "OP20");
            data.OP20_Setup = GetDouble(props, "OP20_S");
            data.OP20_Run = GetDouble(props, "OP20_R");

            data.HasPressbrake = GetString(props, "Pressbrake") == "Checked" ||
                                 GetString(props, "PressBrake") == "Checked";
            data.F140_Setup = GetDouble(props, "F140_S");
            data.F140_Run = GetDouble(props, "F140_R");

            data.F210_Setup = GetDouble(props, "F210_S");
            data.F210_Run = GetDouble(props, "F210_R");

            data.F220_Setup = GetDouble(props, "F220_S");
            data.F220_Run = GetDouble(props, "F220_R");

            data.F325_Setup = GetDouble(props, "F325_S");
            data.F325_Run = GetDouble(props, "F325_R");

            data.Other_Setup = GetDouble(props, "Other_S") + GetDouble(props, "Other_S2");
            data.Other_Run = GetDouble(props, "Other_R") + GetDouble(props, "Other_R2");

            // Determine part type
            string rbPartType = GetString(props, "rbPartType");
            string rbPartTypeSub = GetString(props, "rbPartTypeSub");

            if (rbPartType == "1")
            {
                if (rbPartTypeSub == "1")
                    data.PartType = "Purchased";
                else if (rbPartTypeSub == "2")
                    data.PartType = "Customer-Supplied";
                else
                    data.PartType = "Machined";
            }
            else
            {
                data.PartType = "Manufactured";
            }

            return data;
        }

        private static string GetString(Dictionary<string, string> props, string key)
        {
            return props.TryGetValue(key, out var val) ? val ?? "" : "";
        }

        private static double GetDouble(Dictionary<string, string> props, string key)
        {
            if (props.TryGetValue(key, out var val) && !string.IsNullOrEmpty(val))
            {
                if (double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                    return d;
            }
            return 0;
        }
    }
}
