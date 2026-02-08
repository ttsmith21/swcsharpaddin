using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using NM.Core.DataModel;

namespace NM.Core.Export
{
    /// <summary>
    /// Generates Excel (.xlsx) quote reports from processed part data.
    /// Uses ClosedXML (no Excel installation required).
    /// </summary>
    public sealed class QuoteExcelGenerator
    {
        /// <summary>
        /// One row in the quote report, mapped from pipeline PartData.
        /// </summary>
        public sealed class QuoteLineItem
        {
            public string PartName { get; set; }
            public string FilePath { get; set; }
            public string Classification { get; set; }
            public string OptiMaterial { get; set; }
            public double MaterialWeight_lb { get; set; }
            public int Quantity { get; set; } = 1;

            // OP20 (Laser/Waterjet)
            public string OP20_WorkCenter { get; set; }
            public double OP20_Setup_hr { get; set; }
            public double OP20_Run_hr { get; set; }
            public double OP20_Cost { get; set; }

            // F140 (Press Brake)
            public bool HasPressbrake { get; set; }
            public double F140_Setup_hr { get; set; }
            public double F140_Run_hr { get; set; }
            public double F140_Cost { get; set; }

            // Other work centers
            public double F210_Cost { get; set; }
            public double F220_Cost { get; set; }
            public double F325_Cost { get; set; }

            // Totals
            public double TotalMaterialCost { get; set; }
            public double TotalProcessingCost { get; set; }
            public double UnitCost { get; set; }
            public double ExtendedCost { get; set; }
        }

        /// <summary>
        /// Header metadata for the quote report.
        /// </summary>
        public sealed class QuoteReportMetadata
        {
            public string AssemblyName { get; set; }
            public string Customer { get; set; }
            public string SourcePath { get; set; }
            public DateTime GeneratedAt { get; set; } = DateTime.Now;
        }

        /// <summary>
        /// Maps pipeline PartData + quantity to a QuoteLineItem.
        /// </summary>
        public static QuoteLineItem FromPartData(PartData pd, int quantity)
        {
            if (pd == null) throw new ArgumentNullException(nameof(pd));

            var cost = pd.Cost ?? new CostingData();
            var item = new QuoteLineItem
            {
                PartName = pd.PartName ?? Path.GetFileNameWithoutExtension(pd.FilePath ?? ""),
                FilePath = pd.FilePath,
                Classification = pd.Classification.ToString(),
                OptiMaterial = pd.OptiMaterial ?? pd.Material ?? "",
                MaterialWeight_lb = cost.MaterialWeight_lb,
                Quantity = quantity,

                OP20_WorkCenter = cost.OP20_WorkCenter ?? "",
                OP20_Setup_hr = cost.OP20_S_min / 60.0,
                OP20_Run_hr = cost.OP20_R_min / 60.0,
                OP20_Cost = cost.F115_Price,

                HasPressbrake = pd.Sheet.BendCount > 0,
                F140_Setup_hr = cost.F140_S_min / 60.0,
                F140_Run_hr = cost.F140_R_min / 60.0,
                F140_Cost = cost.F140_Price,

                F210_Cost = cost.F210_Price,
                F220_Cost = cost.F220_Price,
                F325_Cost = cost.F325_Price,

                TotalMaterialCost = cost.TotalMaterialCost,
                TotalProcessingCost = cost.TotalProcessingCost,
                UnitCost = cost.TotalCost,
                ExtendedCost = cost.TotalCost * quantity
            };

            return item;
        }

        /// <summary>
        /// Generates an Excel quote report and saves to the specified path.
        /// </summary>
        public void Generate(List<QuoteLineItem> items, QuoteReportMetadata metadata, string outputPath)
        {
            if (items == null) throw new ArgumentNullException(nameof(items));
            if (string.IsNullOrEmpty(outputPath)) throw new ArgumentNullException(nameof(outputPath));
            metadata = metadata ?? new QuoteReportMetadata();

            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("Quote");

                // --- Header section (rows 1-4) ---
                WriteHeader(ws, metadata);

                // --- Column headers (row 6) ---
                int headerRow = 6;
                WriteColumnHeaders(ws, headerRow);

                // --- Data rows (starting row 7) ---
                int dataStartRow = headerRow + 1;
                for (int i = 0; i < items.Count; i++)
                {
                    WriteDataRow(ws, dataStartRow + i, items[i], i % 2 == 1);
                }

                // --- Summary row ---
                int summaryRow = dataStartRow + items.Count;
                WriteSummaryRow(ws, headerRow, dataStartRow, summaryRow, items.Count);

                // --- Formatting ---
                ApplyFormatting(ws, headerRow, dataStartRow, summaryRow);

                // Save
                var dir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                workbook.SaveAs(outputPath);
            }
        }

        private static void WriteHeader(IXLWorksheet ws, QuoteReportMetadata metadata)
        {
            ws.Cell("A1").Value = "Quote Report";
            ws.Cell("A1").Style.Font.Bold = true;
            ws.Cell("A1").Style.Font.FontSize = 14;

            ws.Cell("A2").Value = "Assembly:";
            ws.Cell("B2").Value = metadata.AssemblyName ?? "";
            ws.Cell("A2").Style.Font.Bold = true;

            ws.Cell("A3").Value = "Customer:";
            ws.Cell("B3").Value = metadata.Customer ?? "";
            ws.Cell("A3").Style.Font.Bold = true;

            ws.Cell("A4").Value = "Generated:";
            ws.Cell("B4").Value = metadata.GeneratedAt.ToString("yyyy-MM-dd HH:mm");
            ws.Cell("A4").Style.Font.Bold = true;
        }

        private static readonly string[] ColumnHeaders = new[]
        {
            "Part Name",     // A
            "Type",          // B
            "Material",      // C
            "Weight (lb)",   // D
            "Qty",           // E
            "OP20",          // F
            "OP20 Setup",    // G
            "OP20 Run",      // H
            "OP20 Cost",     // I
            "Brake?",        // J
            "F140 Setup",    // K
            "F140 Run",      // L
            "F140 Cost",     // M
            "F210 Cost",     // N
            "F220 Cost",     // O
            "F325 Cost",     // P
            "Material $",    // Q
            "Processing $",  // R
            "Unit Cost",     // S
            "Extended $"     // T
        };

        private static void WriteColumnHeaders(IXLWorksheet ws, int row)
        {
            for (int col = 1; col <= ColumnHeaders.Length; col++)
            {
                var cell = ws.Cell(row, col);
                cell.Value = ColumnHeaders[col - 1];
                cell.Style.Font.Bold = true;
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(70, 130, 180); // Steel blue
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
        }

        private static void WriteDataRow(IXLWorksheet ws, int row, QuoteLineItem item, bool shaded)
        {
            // A: Part Name (hyperlink if file exists)
            var nameCell = ws.Cell(row, 1);
            nameCell.Value = item.PartName ?? "";
            if (!string.IsNullOrEmpty(item.FilePath) && File.Exists(item.FilePath))
            {
                try
                {
                    nameCell.SetHyperlink(new XLHyperlink(item.FilePath));
                    nameCell.Style.Font.FontColor = XLColor.Blue;
                    nameCell.Style.Font.Underline = XLFontUnderlineValues.Single;
                }
                catch
                {
                    // Hyperlink creation can fail for unusual paths - skip silently
                }
            }

            // B-E: Classification, Material, Weight, Qty
            ws.Cell(row, 2).Value = item.Classification ?? "";
            ws.Cell(row, 3).Value = item.OptiMaterial ?? "";
            ws.Cell(row, 4).Value = item.MaterialWeight_lb;
            ws.Cell(row, 5).Value = item.Quantity;

            // F-I: OP20
            ws.Cell(row, 6).Value = item.OP20_WorkCenter ?? "";
            ws.Cell(row, 7).Value = item.OP20_Setup_hr;
            ws.Cell(row, 8).Value = item.OP20_Run_hr;
            ws.Cell(row, 9).Value = item.OP20_Cost;

            // J-M: Press Brake
            ws.Cell(row, 10).Value = item.HasPressbrake ? "Y" : "N";
            ws.Cell(row, 11).Value = item.F140_Setup_hr;
            ws.Cell(row, 12).Value = item.F140_Run_hr;
            ws.Cell(row, 13).Value = item.F140_Cost;

            // N-P: Other work centers
            ws.Cell(row, 14).Value = item.F210_Cost;
            ws.Cell(row, 15).Value = item.F220_Cost;
            ws.Cell(row, 16).Value = item.F325_Cost;

            // Q-T: Cost totals
            ws.Cell(row, 17).Value = item.TotalMaterialCost;
            ws.Cell(row, 18).Value = item.TotalProcessingCost;
            ws.Cell(row, 19).Value = item.UnitCost;
            ws.Cell(row, 20).Value = item.ExtendedCost;

            // Alternating row shading
            if (shaded)
            {
                var range = ws.Range(row, 1, row, 20);
                range.Style.Fill.BackgroundColor = XLColor.FromArgb(240, 240, 240);
            }
        }

        private static void WriteSummaryRow(IXLWorksheet ws, int headerRow, int dataStart, int summaryRow, int itemCount)
        {
            if (itemCount == 0) return;

            int dataEnd = dataStart + itemCount - 1;

            ws.Cell(summaryRow, 1).Value = "TOTALS";
            ws.Cell(summaryRow, 1).Style.Font.Bold = true;

            // SUM formulas for numeric columns
            // D: Weight
            ws.Cell(summaryRow, 4).FormulaA1 = $"SUM(D{dataStart}:D{dataEnd})";
            // E: Qty
            ws.Cell(summaryRow, 5).FormulaA1 = $"SUM(E{dataStart}:E{dataEnd})";
            // I: OP20 Cost
            ws.Cell(summaryRow, 9).FormulaA1 = $"SUM(I{dataStart}:I{dataEnd})";
            // M: F140 Cost
            ws.Cell(summaryRow, 13).FormulaA1 = $"SUM(M{dataStart}:M{dataEnd})";
            // N: F210 Cost
            ws.Cell(summaryRow, 14).FormulaA1 = $"SUM(N{dataStart}:N{dataEnd})";
            // O: F220 Cost
            ws.Cell(summaryRow, 15).FormulaA1 = $"SUM(O{dataStart}:O{dataEnd})";
            // P: F325 Cost
            ws.Cell(summaryRow, 16).FormulaA1 = $"SUM(P{dataStart}:P{dataEnd})";
            // Q: Material $
            ws.Cell(summaryRow, 17).FormulaA1 = $"SUM(Q{dataStart}:Q{dataEnd})";
            // R: Processing $
            ws.Cell(summaryRow, 18).FormulaA1 = $"SUM(R{dataStart}:R{dataEnd})";
            // S: Unit Cost
            ws.Cell(summaryRow, 19).FormulaA1 = $"SUM(S{dataStart}:S{dataEnd})";
            // T: Extended $
            ws.Cell(summaryRow, 20).FormulaA1 = $"SUM(T{dataStart}:T{dataEnd})";

            // Bold + top border for summary row
            var range = ws.Range(summaryRow, 1, summaryRow, 20);
            range.Style.Font.Bold = true;
            range.Style.Border.TopBorder = XLBorderStyleValues.Medium;
        }

        private static void ApplyFormatting(IXLWorksheet ws, int headerRow, int dataStartRow, int summaryRow)
        {
            // Number formats for data + summary rows
            int lastRow = summaryRow;

            // D: Weight - 0.00
            ws.Range(dataStartRow, 4, lastRow, 4).Style.NumberFormat.Format = "0.00";
            // G, H: OP20 Setup/Run hours - 0.0000
            ws.Range(dataStartRow, 7, lastRow, 8).Style.NumberFormat.Format = "0.0000";
            // K, L: F140 Setup/Run hours - 0.0000
            ws.Range(dataStartRow, 11, lastRow, 12).Style.NumberFormat.Format = "0.0000";
            // I, M-T: Currency columns - $#,##0.00
            ws.Range(dataStartRow, 9, lastRow, 9).Style.NumberFormat.Format = "$#,##0.00";
            ws.Range(dataStartRow, 13, lastRow, 20).Style.NumberFormat.Format = "$#,##0.00";

            // Freeze panes below header row
            ws.SheetView.FreezeRows(headerRow);

            // Auto-fit columns
            ws.Columns(1, 20).AdjustToContents();

            // Set minimum width for narrow columns
            for (int col = 1; col <= 20; col++)
            {
                if (ws.Column(col).Width < 8)
                    ws.Column(col).Width = 8;
            }
        }
    }
}
