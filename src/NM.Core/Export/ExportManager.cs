using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using NM.Core.DataModel;

namespace NM.Core.Export
{
    // Minimal exporter to unblock quoting/ERP without adding Excel COM/OpenXML.
    // CSV opens in Excel; ERP export is a simple delimited text.
    public sealed class ExportManager
    {
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        private const double M_TO_IN = 39.37007874015748;
        private const double KG_TO_LB = 2.2046226218487757;

        public void ExportToCsv(IEnumerable<PartData> parts, string csvPath)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(csvPath) ?? ".");
            using (var sw = new StreamWriter(csvPath, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true)))
            {
                sw.WriteLine(string.Join(",",
                    "PartName","Configuration","ParentAssembly","Status","Classification",
                    "IsSheetMetal","IsTube","Thickness_in","RawWeight_lb",
                    "OP20_S","OP20_R","F115_Price",
                    "F140_S","F140_R","F140_S_Cost","F140_Price",
                    "F220","F220_S","F220_R","F220_RN",
                    "QuoteQty","TotalPrice","FailureReason"));

                foreach (var p in parts ?? Enumerable.Empty<PartData>())
                {
                    var thickness_in = p.Thickness_m * M_TO_IN;
                    var rawWeight_lb = p.Mass_kg * KG_TO_LB;

                    sw.WriteLine(string.Join(",",
                        Csv(p.PartName),
                        Csv(p.Configuration),
                        Csv(p.ParentAssembly),
                        Csv(p.Status.ToString()),
                        Csv(p.Classification.ToString()),
                        Csv(p.Sheet.IsSheetMetal),
                        Csv(p.Tube.IsTube),
                        Csv(thickness_in),
                        Csv(rawWeight_lb),
                        Csv(p.Cost.OP20_S_min),
                        Csv(p.Cost.OP20_R_min),
                        Csv(p.Cost.F115_Price),
                        Csv(p.Cost.F140_S_min),
                        Csv(p.Cost.F140_R_min),
                        Csv(p.Cost.F140_S_Cost),
                        Csv(p.Cost.F140_Price),
                        Csv(p.Cost.F220_min),
                        Csv(p.Cost.F220_S_min),
                        Csv(p.Cost.F220_R_min),
                        Csv(p.Cost.F220_RN),
                        Csv(p.QuoteQty),
                        Csv(p.TotalPrice),
                        Csv(p.FailureReason)
                    ));
                }
            }
        }

        public void ExportToErp(IEnumerable<PartData> parts, string txtPath, char delimiter = '\t')
        {
            Directory.CreateDirectory(Path.GetDirectoryName(txtPath) ?? ".");
            using (var sw = new StreamWriter(txtPath, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true)))
            {
                // Minimal ERP routing: PartName, ParentAssembly, OptiMaterial, Qty, TotalPrice
                sw.WriteLine(string.Join(delimiter.ToString(), "PartName","ParentAssembly","OptiMaterial","Qty","TotalPrice"));
                foreach (var p in parts ?? Enumerable.Empty<PartData>())
                {
                    sw.WriteLine(string.Join(delimiter.ToString(),
                        Safe(p.PartName),
                        Safe(p.ParentAssembly),
                        Safe(p.OptiMaterial),
                        p.QuoteQty.ToString(Inv),
                        p.TotalPrice.ToString("0.####", Inv)
                    ));
                }
            }
        }

        private static string Csv(string s)
        {
            if (s == null) return "";
            var needsQuotes = s.Contains(",") || s.Contains("\"") || s.Contains("\n") || s.Contains("\r");
            if (!needsQuotes) return s;
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        }

        private static string Csv(bool b) => b ? "True" : "False";
        private static string Csv(int i) => i.ToString(Inv);
        private static string Csv(double d) => d.ToString("0.####", Inv);
        private static string Safe(string s) => s ?? "";
    }
}
