using System;
using System.Collections.Generic;
using System.Globalization;
using NM.Core.DataModel;

namespace NM.Core.Processing
{
    // Converts between strongly-typed PartData and the string-based custom properties schema.
    // Centralizes string conversions and legacy property naming.
    public static class PartDataPropertyMap
    {
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        private const double M_TO_IN = 39.37007874015748;
        private const double KG_TO_LB = 2.2046226218487757;

        public static IDictionary<string, string> ToProperties(PartData d)
        {
            var p = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var thickness_in = d.Thickness_m * M_TO_IN;
            var rawWeight_lb = d.Mass_kg * KG_TO_LB;

            p["RawWeight"] = rawWeight_lb.ToString("0.####", Inv);
            p["SheetPercent"] = d.SheetPercent.ToString("0.####", Inv);

            // OP20 / Laser
            p["OP20_S"] = d.Cost.OP20_S_min.ToString("0.####", Inv);
            p["OP20_R"] = d.Cost.OP20_R_min.ToString("0.####", Inv);
            p["F115_Price"] = d.Cost.F115_Price.ToString("0.####", Inv);

            // F140
            p["F140_S"] = d.Cost.F140_S_min.ToString("0.####", Inv);
            p["F140_R"] = d.Cost.F140_R_min.ToString("0.####", Inv);
            p["F140_S_Cost"] = d.Cost.F140_S_Cost.ToString("0.####", Inv);
            p["F140_Price"] = d.Cost.F140_Price.ToString("0.####", Inv);

            // F210 Deburr
            p["F210_S"] = d.Cost.F210_S_min.ToString("0.####", Inv);
            p["F210_R"] = d.Cost.F210_R_min.ToString("0.####", Inv);
            p["F210_Price"] = d.Cost.F210_Price.ToString("0.####", Inv);

            // F220 Tapping
            p["F220"] = d.Cost.F220_min.ToString("0.####", Inv);
            p["F220_S"] = d.Cost.F220_S_min.ToString("0.####", Inv);
            p["F220_R"] = d.Cost.F220_R_min.ToString("0.####", Inv);
            p["F220_RN"] = d.Cost.F220_RN.ToString(Inv);
            p["F220_Note"] = d.Cost.F220_Note ?? string.Empty;
            p["F220_Price"] = d.Cost.F220_Price.ToString("0.####", Inv);

            // F325 Roll Forming
            p["F325_S"] = d.Cost.F325_S_min.ToString("0.####", Inv);
            p["F325_R"] = d.Cost.F325_R_min.ToString("0.####", Inv);
            p["F325_Price"] = d.Cost.F325_Price.ToString("0.####", Inv);

            // Material costs and totals
            p["MaterialCost"] = d.Cost.MaterialCost.ToString("0.##", Inv);
            p["MaterialWeight_lb"] = d.Cost.MaterialWeight_lb.ToString("0.####", Inv);
            p["TotalMaterialCost"] = d.Cost.TotalMaterialCost.ToString("0.##", Inv);
            p["TotalProcessingCost"] = d.Cost.TotalProcessingCost.ToString("0.##", Inv);
            p["TotalCost"] = d.Cost.TotalCost.ToString("0.##", Inv);
            if (!string.IsNullOrEmpty(d.Cost.OP20_WorkCenter))
                p["OP20_WorkCenter"] = d.Cost.OP20_WorkCenter;

            // Material and basics
            p["MaterialCostPerLB"] = d.MaterialCostPerLB.ToString("0.####", Inv);
            p["MaterailCostPerLB"] = d.MaterialCostPerLB.ToString("0.####", Inv); // legacy misspelling
            p["QuoteQty"] = d.QuoteQty.ToString(Inv);
            p["TotalPrice"] = d.TotalPrice.ToString("0.####", Inv);

            p["OptiMaterial"] = d.OptiMaterial ?? string.Empty;
            p["rbMaterialType"] = d.Material ?? string.Empty;
            p["MaterialCategory"] = d.MaterialCategory ?? string.Empty;

            p["Thickness"] = thickness_in.ToString("0.####", Inv);
            p["IsSheetMetal"] = d.Sheet.IsSheetMetal ? "True" : "False";
            p["IsTube"] = d.Tube.IsTube ? "True" : "False";

            foreach (var kv in d.Extra)
                p[kv.Key] = kv.Value ?? string.Empty;

            return p;
        }

        // Optional helper when backfilling DTOs from existing files (vNext).
        public static PartData FromProperties(IDictionary<string, string> props)
        {
            var d = new PartData();

            double D(string k) => TryD(props, k, out var x) ? x : 0.0;
            int I(string k) => TryI(props, k, out var x) ? x : 0;

            d.Mass_kg = D("RawWeight") / KG_TO_LB; // RawWeight was in lb
            d.SheetPercent = D("SheetPercent");

            d.Cost.OP20_S_min = D("OP20_S");
            d.Cost.OP20_R_min = D("OP20_R");
            d.Cost.F115_Price = D("F115_Price");

            d.Cost.F140_S_min = D("F140_S");
            d.Cost.F140_R_min = D("F140_R");
            d.Cost.F140_S_Cost = D("F140_S_Cost");
            d.Cost.F140_Price = D("F140_Price");

            // F210 Deburr
            d.Cost.F210_S_min = D("F210_S");
            d.Cost.F210_R_min = D("F210_R");
            d.Cost.F210_Price = D("F210_Price");

            // F220 Tapping
            d.Cost.F220_min = D("F220");
            d.Cost.F220_S_min = D("F220_S");
            d.Cost.F220_R_min = D("F220_R");
            d.Cost.F220_RN = I("F220_RN");
            d.Cost.F220_Note = Get(props, "F220_Note");
            d.Cost.F220_Price = D("F220_Price");

            // F325 Roll Forming
            d.Cost.F325_S_min = D("F325_S");
            d.Cost.F325_R_min = D("F325_R");
            d.Cost.F325_Price = D("F325_Price");

            // Material costs and totals
            d.Cost.MaterialCost = D("MaterialCost");
            d.Cost.MaterialWeight_lb = D("MaterialWeight_lb");
            d.Cost.TotalMaterialCost = D("TotalMaterialCost");
            d.Cost.TotalProcessingCost = D("TotalProcessingCost");
            d.Cost.TotalCost = D("TotalCost");
            d.Cost.OP20_WorkCenter = Get(props, "OP20_WorkCenter");

            d.MaterialCostPerLB = D("MaterialCostPerLB");
            if (d.MaterialCostPerLB == 0)
                d.MaterialCostPerLB = D("MaterailCostPerLB");

            d.QuoteQty = I("QuoteQty");
            d.TotalPrice = D("TotalPrice");

            d.OptiMaterial = Get(props, "OptiMaterial");
            d.Material = Get(props, "rbMaterialType");
            d.MaterialCategory = Get(props, "MaterialCategory");

            var thickness_in = D("Thickness");
            d.Thickness_m = thickness_in / M_TO_IN;

            d.Sheet.IsSheetMetal = EqualsTrue(Get(props, "IsSheetMetal"));
            d.Tube.IsTube = EqualsTrue(Get(props, "IsTube"));

            return d;
        }

        private static bool TryD(IDictionary<string, string> p, string k, out double v)
            => double.TryParse(Get(p, k), NumberStyles.Float, Inv, out v);
        private static bool TryI(IDictionary<string, string> p, string k, out int v)
            => int.TryParse(Get(p, k), NumberStyles.Integer, Inv, out v);
        private static string Get(IDictionary<string, string> p, string k)
            => p != null && p.TryGetValue(k, out var v) ? v ?? string.Empty : string.Empty;
        private static bool EqualsTrue(string s)
            => "true".Equals(s?.Trim(), StringComparison.OrdinalIgnoreCase);
    }
}
