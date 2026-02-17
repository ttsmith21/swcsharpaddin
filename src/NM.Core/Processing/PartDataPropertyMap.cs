using System;
using System.Collections.Generic;
using System.Globalization;
using NM.Core.DataModel;
using static NM.Core.Constants.UnitConversions;

namespace NM.Core.Processing
{
    // Converts between strongly-typed PartData and the string-based custom properties schema.
    // Centralizes string conversions and legacy property naming.
    //
    // IMPORTANT: All setup/run time properties are written in HOURS (not minutes).
    // Internally CostingData stores times in minutes (*_min fields), but VBA, Tab Builder,
    // and ERP Import.prn all expect HOURS. We convert here: hours = minutes / 60.
    public static class PartDataPropertyMap
    {
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        /// <summary>
        /// Maps internal work center codes to the full OP20 property strings
        /// that match the Tab Builder ComboBox entries (from OP20.txt).
        /// VBA examples: "N120 - 5040", "F110 - TUBE LASER", "F300 - SAW"
        /// </summary>
        private static readonly Dictionary<string, string> OP20Descriptions =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "F115", "N120 - 5040" },         // Default flat laser (VBA default for sheet metal)
            { "N115", "N115 - ANY FLAT LASER" },
            { "N120", "N120 - 5040" },
            { "N125", "N125 - 3060" },
            { "N135", "N135 - FLOW" },
            { "F110", "F110 - TUBE LASER" },
            { "N145", "N145 - 5-AXIS LASER" },
            { "F300", "F300 - SAW" },
            { "NPUR", "NPUR - PURCHASED" },
            { "CUST", "CUST - SUPPLIED" },
            { "MP",   "MP - MACHINED" },
        };

        /// <summary>
        /// Convert minutes to hours, using VBA's format convention.
        /// VBA minimum for setup: 0.01 hours (36 seconds).
        /// </summary>
        private static string MinToHours(double minutes, string format = "0.####")
        {
            double hours = minutes / 60.0;
            return hours.ToString(format, Inv);
        }

        public static IDictionary<string, string> ToProperties(PartData d)
        {
            var p = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var thickness_in = d.Thickness_m * MetersToInches;
            var rawWeight_lb = d.Mass_kg * KgToLbs;

            p["RawWeight"] = rawWeight_lb.ToString("0.####", Inv);
            p["SheetPercent"] = d.SheetPercent.ToString("0.####", Inv);

            // OP20 work center + times (Tab Builder shows OP20 as a ComboBox)
            // Write the full description string (e.g., "N120 - 5040") to match OP20.txt entries.
            if (!string.IsNullOrEmpty(d.Cost.OP20_WorkCenter))
            {
                string op20Desc;
                if (OP20Descriptions.TryGetValue(d.Cost.OP20_WorkCenter, out op20Desc))
                    p["OP20"] = op20Desc;
                else
                    p["OP20"] = d.Cost.OP20_WorkCenter; // fallback to raw code if unknown
                p["OP20_WorkCenter"] = d.Cost.OP20_WorkCenter;

                // Only write OP20 times when we actually computed them (have a work center).
                // Prevents overwriting good values from a previous run with zeros.
                p["OP20_S"] = MinToHours(d.Cost.OP20_S_min);
                p["OP20_R"] = MinToHours(d.Cost.OP20_R_min);
                p["F115_Price"] = d.Cost.F115_Price.ToString("0.####", Inv);
            }

            // F210 Deburr — only write when calculated (prevents zeroing good values on re-run)
            bool f210Active = d.Cost.F210_S_min > 0 || d.Cost.F210_R_min > 0;
            if (f210Active)
            {
                p["F210"] = "1";
                p["F210_S"] = MinToHours(d.Cost.F210_S_min);
                p["F210_R"] = MinToHours(d.Cost.F210_R_min);
                p["F210_Price"] = d.Cost.F210_Price.ToString("0.####", Inv);
            }

            // F220 Tapping — only write when calculated (prevents zeroing on re-run)
            bool f220Active = d.Cost.F220_S_min > 0 || d.Cost.F220_R_min > 0 || d.Cost.F220_RN > 0;
            if (f220Active)
            {
                p["F220"] = "1";
                p["F220_S"] = MinToHours(d.Cost.F220_S_min);
                p["F220_R"] = MinToHours(d.Cost.F220_R_min);
                p["F220_RN"] = d.Cost.F220_RN.ToString(Inv);
                p["F220_Note"] = d.Cost.F220_Note ?? string.Empty;
                p["F220_Price"] = d.Cost.F220_Price.ToString("0.####", Inv);
                if (d.Cost.F220_RN > 0)
                    p["TappedHoleCount"] = d.Cost.F220_RN.ToString(Inv);
            }

            // F140 Press Brake — only write when calculated (prevents zeroing on re-run)
            // VBA: PressBrake="Checked" when sheet metal has bends, or heavy tube needs brake
            bool pressBrakeActive = d.Cost.F140_S_min > 0 || d.Cost.F140_R_min > 0
                                  || d.Sheet.BendCount > 0;
            if (pressBrakeActive)
            {
                p["PressBrake"] = "Checked";
                p["F140_S"] = MinToHours(d.Cost.F140_S_min);
                p["F140_R"] = MinToHours(d.Cost.F140_R_min);
                p["F140_S_Cost"] = d.Cost.F140_S_Cost.ToString("0.####", Inv);
                p["F140_Price"] = d.Cost.F140_Price.ToString("0.####", Inv);
                if (d.Sheet.BendCount > 0)
                    p["BendCount"] = d.Sheet.BendCount.ToString(Inv);
            }

            // F325 Roll Forming — only write when calculated (prevents zeroing on re-run)
            bool f325Active = d.Cost.F325_S_min > 0 || d.Cost.F325_R_min > 0;
            if (f325Active)
            {
                p["F325"] = "1";
                p["F325_S"] = MinToHours(d.Cost.F325_S_min);
                p["F325_R"] = MinToHours(d.Cost.F325_R_min);
                p["F325_Price"] = d.Cost.F325_Price.ToString("0.####", Inv);
            }

            // Material costs and totals
            p["MaterialCost"] = d.Cost.MaterialCost.ToString("0.##", Inv);
            p["MaterialWeight_lb"] = d.Cost.MaterialWeight_lb.ToString("0.####", Inv);
            p["TotalMaterialCost"] = d.Cost.TotalMaterialCost.ToString("0.##", Inv);
            p["TotalProcessingCost"] = d.Cost.TotalProcessingCost.ToString("0.##", Inv);
            p["TotalCost"] = d.Cost.TotalCost.ToString("0.##", Inv);

            // Material and basics
            p["MaterialCostPerLB"] = d.MaterialCostPerLB.ToString("0.####", Inv);
            p["MaterailCostPerLB"] = d.MaterialCostPerLB.ToString("0.####", Inv); // legacy misspelling
            p["QuoteQty"] = d.QuoteQty.ToString(Inv);
            p["TotalPrice"] = d.TotalPrice.ToString("0.####", Inv);

            p["OptiMaterial"] = d.OptiMaterial ?? string.Empty;

            // rbMaterialType is a Tab Builder RadioButton: "0"=Sheet, "1"=Tube, "2"=SqFt
            // VBA writes "1" for tubes (SP.bas:2517,2543); sheet metal defaults to "0".
            if (d.Tube.IsTube)
                p["rbMaterialType"] = "1";
            else
                p["rbMaterialType"] = "0";

            p["MaterialCategory"] = d.MaterialCategory ?? string.Empty;

            p["Thickness"] = thickness_in.ToString("0.####", Inv);
            p["IsSheetMetal"] = d.Sheet.IsSheetMetal ? "True" : "False";
            p["IsTube"] = d.Tube.IsTube ? "True" : "False";

            // Tube-specific properties (dimensions in inches)
            if (d.Tube.IsTube)
            {
                p["TubeOD"] = (d.Tube.OD_m * MetersToInches).ToString("0.####", Inv);
                p["TubeWall"] = (d.Tube.Wall_m * MetersToInches).ToString("0.####", Inv);
                p["TubeLength"] = (d.Tube.Length_m * MetersToInches).ToString("0.####", Inv);
                p["TubeShape"] = d.Tube.TubeShape ?? "Round";
                if (!string.IsNullOrEmpty(d.Tube.NpsText))
                    p["TubeNPS"] = d.Tube.NpsText;
                if (!string.IsNullOrEmpty(d.Tube.ScheduleCode))
                    p["TubeSchedule"] = d.Tube.ScheduleCode;
                if (d.Tube.NumberOfHoles > 0)
                    p["NumberOfHoles"] = d.Tube.NumberOfHoles.ToString(Inv);
                // Tube cut length for Bar/Pipe/Tube weight calc
                if (d.Tube.CutLength_m > 0)
                    p["F300_Length"] = (d.Tube.CutLength_m * MetersToInches).ToString("0.####", Inv);
            }

            // WPS / Welding properties — only write when WPS resolution was attempted
            if (d.Welding.WasResolved)
            {
                p["WPS_Number"] = d.Welding.WpsNumber ?? string.Empty;
                p["WPS_Process"] = d.Welding.WeldProcess ?? string.Empty;
                p["WPS_FillerMetal"] = d.Welding.FillerMetal ?? string.Empty;
                p["WPS_JointType"] = d.Welding.JointType ?? string.Empty;
                p["WPS_NeedsReview"] = d.Welding.NeedsReview ? "True" : "False";
                if (d.Welding.NeedsReview)
                    p["WPS_ReviewReasons"] = d.Welding.ReviewReasons ?? string.Empty;
            }

            // Extras (ERP props, user-entered values, etc.) — written last so they can override
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

            d.Mass_kg = D("RawWeight") / KgToLbs; // RawWeight was in lb
            d.SheetPercent = D("SheetPercent");

            // Properties store times in HOURS; internal storage is minutes → multiply by 60
            d.Cost.OP20_S_min = D("OP20_S") * 60.0;
            d.Cost.OP20_R_min = D("OP20_R") * 60.0;
            d.Cost.F115_Price = D("F115_Price");

            d.Cost.F140_S_min = D("F140_S") * 60.0;
            d.Cost.F140_R_min = D("F140_R") * 60.0;
            d.Cost.F140_S_Cost = D("F140_S_Cost");
            d.Cost.F140_Price = D("F140_Price");

            // F210 Deburr
            d.Cost.F210_S_min = D("F210_S") * 60.0;
            d.Cost.F210_R_min = D("F210_R") * 60.0;
            d.Cost.F210_Price = D("F210_Price");

            // F220 Tapping
            d.Cost.F220_min = D("F220");
            d.Cost.F220_S_min = D("F220_S") * 60.0;
            d.Cost.F220_R_min = D("F220_R") * 60.0;
            d.Cost.F220_RN = I("F220_RN");
            d.Cost.F220_Note = Get(props, "F220_Note");
            d.Cost.F220_Price = D("F220_Price");

            // F325 Roll Forming
            d.Cost.F325_S_min = D("F325_S") * 60.0;
            d.Cost.F325_R_min = D("F325_R") * 60.0;
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
            // rbMaterialType is "0"/"1"/"2" (radio button), NOT a material name.
            // d.Material should come from the SolidWorks material assignment, not from this property.
            d.MaterialCategory = Get(props, "MaterialCategory");

            var thickness_in = D("Thickness");
            d.Thickness_m = thickness_in / MetersToInches;

            d.Sheet.IsSheetMetal = EqualsTrue(Get(props, "IsSheetMetal"));
            d.Tube.IsTube = EqualsTrue(Get(props, "IsTube"));

            // Tube-specific properties
            if (d.Tube.IsTube)
            {
                d.Tube.OD_m = D("TubeOD") / MetersToInches;
                d.Tube.Wall_m = D("TubeWall") / MetersToInches;
                d.Tube.Length_m = D("TubeLength") / MetersToInches;
                d.Tube.TubeShape = Get(props, "TubeShape");
                d.Tube.NpsText = Get(props, "TubeNPS");
                d.Tube.ScheduleCode = Get(props, "TubeSchedule");
                d.Tube.NumberOfHoles = I("NumberOfHoles");
            }

            // WPS / Welding properties
            var wpsNum = Get(props, "WPS_Number");
            if (!string.IsNullOrEmpty(wpsNum))
            {
                d.Welding.WasResolved = true;
                d.Welding.WpsNumber = wpsNum;
                d.Welding.WeldProcess = Get(props, "WPS_Process");
                d.Welding.FillerMetal = Get(props, "WPS_FillerMetal");
                d.Welding.JointType = Get(props, "WPS_JointType");
                d.Welding.NeedsReview = EqualsTrue(Get(props, "WPS_NeedsReview"));
                d.Welding.ReviewReasons = Get(props, "WPS_ReviewReasons");
            }

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
