using System;
using System.Collections.Generic;
using System.Globalization;
using NM.Core;
using NM.Core.Tubes;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin
{
    public sealed class SimpleTubeProcessor
    {
        private readonly ISldWorks _swApp;
        private readonly TubeCuttingParameterService _cutSvc = new TubeCuttingParameterService();
        private readonly PipeScheduleService _pipeSvc = new PipeScheduleService();

        private static readonly HashSet<double> StandardRoundBarSizes = new HashSet<double>
        {
            0.125, 0.1563, 0.1875, 0.25, 0.3125, 0.375, 0.4375, 0.5, 0.625, 0.75,
            0.875, 0.9375, 1.0, 1.1875, 1.25, 1.375, 1.5, 1.75, 2.0, 2.5, 3.0
        };

        public SimpleTubeProcessor(ISldWorks swApp)
        {
            _swApp = swApp ?? throw new ArgumentNullException(nameof(swApp));
        }

        public bool Process(ModelInfo info, IModelDoc2 doc, ProcessingOptions options)
        {
            const string proc = nameof(SimpleTubeProcessor) + "." + nameof(Process);
            ErrorHandler.PushCallStack(proc);
            try
            {
                if (info == null || doc == null) { ErrorHandler.HandleError(proc, "Null model/doc"); return false; }

                string shape = ReadString(info, "Shape");
                string crossSection = ReadString(info, "CrossSection");
                string wallStr = ReadString(info, "Wall Thickness");
                string matLenStr = ReadString(info, "Material Length");
                string cutLenStr = ReadString(info, "Cut Length");
                string holesStr = ReadString(info, "Number of Holes");

                if (string.IsNullOrWhiteSpace(shape) || string.IsNullOrWhiteSpace(crossSection) || string.IsNullOrWhiteSpace(wallStr))
                {
                    ErrorHandler.HandleError(proc, "Missing tube properties (Shape/CrossSection/Wall Thickness)", null, "Warning");
                    return false;
                }

                bool isRound = shape.Trim().Equals("Round", StringComparison.OrdinalIgnoreCase);

                if (!TryParseDoubleInvariant(wallStr, out double wallIn)) wallIn = 0.0;
                if (!TryParseDoubleInvariant(matLenStr, out double materialLengthIn)) materialLengthIn = 0.0;
                if (!TryParseDoubleInvariant(cutLenStr, out double cutLengthIn)) cutLengthIn = materialLengthIn;
                if (!TryParseDoubleInvariant(holesStr, out double holeCount)) holeCount = 0;

                double odIn = TryParseCrossSectionInches(crossSection);

                bool isRoundBar = isRound && RoundBarValidator.IsRoundBar(odIn, odIn - 2.0 * wallIn);

                string material = options?.Material ?? "304L";
                string materialCategory = options?.MaterialCategory.ToString() ?? "StainlessSteel";

                info.CustomProperties.IsTube = true;
                info.IsSheetMetal = false;

                bool processed;
                if (isRound && (odIn > 0))
                {
                    if (isRoundBar || wallIn <= 0 || wallIn > (odIn * 0.3))
                    {
                        processed = ProcessRoundBar(info, odIn, materialLengthIn, material);
                    }
                    else
                    {
                        processed = ProcessRoundTube(info, odIn, wallIn, materialLengthIn, cutLengthIn, holeCount, material, materialCategory);
                    }
                }
                else
                {
                    processed = ProcessNonRound(info, shape, crossSection, wallIn, materialLengthIn, material, materialCategory);
                }

                if (processed)
                {
                    if (readDouble(info, "Weight", out double wtLb) && wtLb > 0)
                    {
                        var f325 = TubeWorkCenterRules.ComputeF325(wtLb, wallIn);
                        info.CustomProperties.SetPropertyValue("F325", f325.Code, CustomPropertyType.Text);
                        info.CustomProperties.SetPropertyValue("F325_R", f325.RunHours.ToString("F3", CultureInfo.InvariantCulture), CustomPropertyType.Number);
                        info.CustomProperties.SetPropertyValue("F325_S", FormatDotText(f325.SetupHours), CustomPropertyType.Text);

                        var f140 = TubeWorkCenterRules.ComputeF140(wtLb, wallIn);
                        if (f325.RequiresPressBrake && f140.RunHours > 0)
                        {
                            info.CustomProperties.SetPropertyValue("F140_S", FormatDotText(f140.SetupHours), CustomPropertyType.Text);
                            info.CustomProperties.SetPropertyValue("F140_R", FormatDotText(f140.RunHours), CustomPropertyType.Text);
                        }
                    }

                    double deburrLen = cutLengthIn > 0 ? cutLengthIn : materialLengthIn;
                    var f210 = TubeWorkCenterRules.ComputeF210(deburrLen);
                    info.CustomProperties.SetPropertyValue("F210_R", f210.RunHours.ToString("F3", CultureInfo.InvariantCulture), CustomPropertyType.Number);
                    info.CustomProperties.SetPropertyValue("F210_S", f210.SetupHours, CustomPropertyType.Number);
                }

                return processed;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Tube processor exception", ex);
                return false;
            }
            finally { ErrorHandler.PopCallStack(); }
        }

        private bool ProcessRoundBar(ModelInfo info, double odIn, double lengthIn, string material)
        {
            info.CustomProperties.SetPropertyValue("OP20", "F300 - SAW", CustomPropertyType.Text);
            info.CustomProperties.SetPropertyValue("rbMaterialType", "1", CustomPropertyType.Text);
            info.CustomProperties.SetPropertyValue("F300_Length", lengthIn.ToString("F3", CultureInfo.InvariantCulture), CustomPropertyType.Number);
            info.CustomProperties.SetPropertyValue("OP20_S", ".05", CustomPropertyType.Text);
            double op20r = ((odIn * 90.0) + 15.0) / 3600.0; // hours
            info.CustomProperties.SetPropertyValue("OP20_R", op20r.ToString("F4", CultureInfo.InvariantCulture), CustomPropertyType.Number);

            if (string.IsNullOrWhiteSpace(ReadString(info, "OptiMaterial")))
            {
                string opti;
                if (IsStandardRoundBar(odIn))
                {
                    opti = $"R.{material}{FormatInchesNoSpace(odIn)}\"";
                }
                else
                {
                    int mm = (int)Math.Floor(odIn * 25.4);
                    opti = $"R.{material}M{mm}";
                }
                info.CustomProperties.SetPropertyValue("OptiMaterial", opti, CustomPropertyType.Text);
            }
            if (string.IsNullOrWhiteSpace(info.CustomProperties.Description))
            {
                info.CustomProperties.SetPropertyValue("Description", $"{material} ROUND", CustomPropertyType.Text);
            }
            return true;
        }

        private static bool IsStandardRoundBar(double odIn)
        {
            const double tol = 0.001;
            foreach (var s in StandardRoundBarSizes)
            {
                if (Math.Abs(odIn - s) <= tol) return true;
            }
            return false;
        }

        private bool ProcessRoundTube(ModelInfo info, double odIn, double wallIn, double lengthIn, double cutLenIn, double holes, string material, string materialCategory)
        {
            string npsText; string schedCode;
            bool isPipe = _pipeSvc.TryResolveByOdAndWall(odIn, wallIn, materialCategory, out npsText, out schedCode);
            string prefix = isPipe ? "P." : "T.";

            string tubeMaterial = material;
            if (material.Equals("A36", StringComparison.OrdinalIgnoreCase) || material.Equals("ALNZD", StringComparison.OrdinalIgnoreCase))
            {
                tubeMaterial = isPipe ? "BLK" : "HR";
            }

            if (string.IsNullOrWhiteSpace(ReadString(info, "OptiMaterial")))
            {
                string opti = isPipe
                    ? $"{prefix}{tubeMaterial}{npsText}SCH{schedCode}"
                    : $"{prefix}{tubeMaterial}{FormatInchesNoSpace(odIn)}\"ODX{FormatDot(wallIn)}\"";
                info.CustomProperties.SetPropertyValue("OptiMaterial", opti, CustomPropertyType.Text);
            }
            if (string.IsNullOrWhiteSpace(info.CustomProperties.Description))
            {
                info.CustomProperties.SetPropertyValue("Description", $"{tubeMaterial} {(isPipe ? "PIPE" : "TUBE")}", CustomPropertyType.Text);
            }

            info.CustomProperties.SetPropertyValue("rbMaterialType", "1", CustomPropertyType.Text);
            info.CustomProperties.SetPropertyValue("F300_Length", lengthIn.ToString("F3", CultureInfo.InvariantCulture), CustomPropertyType.Number);

            if (odIn <= 6.0)
            {
                info.CustomProperties.SetPropertyValue("OP20", "F110 - TUBE LASER", CustomPropertyType.Text);
                info.CustomProperties.SetPropertyValue("OP20_S", ".15", CustomPropertyType.Text);
            }
            else if (odIn <= 10.0)
            {
                info.CustomProperties.SetPropertyValue("OP20", "F110 - TUBE LASER", CustomPropertyType.Text);
                info.CustomProperties.SetPropertyValue("OP20_S", ".5", CustomPropertyType.Text);
            }
            else if (odIn <= 10.75)
            {
                info.CustomProperties.SetPropertyValue("OP20", "F110 - TUBE LASER", CustomPropertyType.Text);
                info.CustomProperties.SetPropertyValue("OP20_S", "1", CustomPropertyType.Text);
            }
            else
            {
                info.CustomProperties.SetPropertyValue("OP20", "N145 - 5-AXIS LASER", CustomPropertyType.Text);
                info.CustomProperties.SetPropertyValue("OP20_S", ".25", CustomPropertyType.Text);
            }

            var p = _cutSvc.Get(materialCategory, wallIn);
            double cutTimeSec = (p.CutSpeedInPerMin > 0 ? (cutLenIn / p.CutSpeedInPerMin) * 60.0 : 0.0);
            double cycleTimeSec = ((lengthIn / 240.0) * 45.0);
            double pierceTimeSec = (holes + 2.0) * (p.PierceTimeSec + 1.5);
            double traverseTimeSec = (lengthIn / 1440.0) * 60.0;
            double totalHours = (cutTimeSec + cycleTimeSec + pierceTimeSec + traverseTimeSec) / 3600.0;

            if (readDouble(info, "Weight", out double wtLb) && wtLb > 50.0 && totalHours < 0.05) totalHours = 0.05;
            if (wallIn > 0.2) totalHours *= 2.0;

            info.CustomProperties.SetPropertyValue("OP20_R", totalHours.ToString("F4", CultureInfo.InvariantCulture), CustomPropertyType.Number);
            return true;
        }

        private bool ProcessNonRound(ModelInfo info, string shape, string cross, double wallIn, double lengthIn, string material, string materialCategory)
        {
            double major = TryParseCrossSectionInches(cross);

            string op; string baseSetup;
            if (major <= 6.0 && major > 0) { op = "F110 - TUBE LASER"; baseSetup = ".15"; }
            else if (major <= 10.0 && major > 0) { op = "F110 - TUBE LASER"; baseSetup = ".5"; }
            else { op = "N145 - 5-AXIS LASER"; baseSetup = ".25"; }

            string tubeMaterial = material;
            if (material.Equals("A36", StringComparison.OrdinalIgnoreCase) || material.Equals("ALNZD", StringComparison.OrdinalIgnoreCase))
            {
                tubeMaterial = "HR";
            }

            info.CustomProperties.SetPropertyValue("rbMaterialType", "1", CustomPropertyType.Text);
            info.CustomProperties.SetPropertyValue("F300_Length", lengthIn.ToString("F3", CultureInfo.InvariantCulture), CustomPropertyType.Number);
            info.CustomProperties.SetPropertyValue("OP20", op, CustomPropertyType.Text);

            double setupVal = 0.0; double.TryParse(baseSetup, NumberStyles.Any, CultureInfo.InvariantCulture, out setupVal);
            bool isAngleOrChannel = shape.Equals("Angle", StringComparison.OrdinalIgnoreCase) || shape.Equals("Channel", StringComparison.OrdinalIgnoreCase);
            if (isAngleOrChannel) setupVal += 0.25;
            info.CustomProperties.SetPropertyValue("OP20_S", FormatDotText(setupVal), CustomPropertyType.Text);

            if (string.IsNullOrWhiteSpace(info.CustomProperties.Description))
            {
                string kind = shape.Equals("Angle", StringComparison.OrdinalIgnoreCase) ? "ANGLE" : "TUBE";
                info.CustomProperties.SetPropertyValue("Description", $"{tubeMaterial} {kind}", CustomPropertyType.Text);
            }
            if (string.IsNullOrWhiteSpace(ReadString(info, "OptiMaterial")))
            {
                // Parse dims for exact formatting
                ParseCrossDims(cross, out var left, out var right);
                string opti;
                if (shape.Equals("Square", StringComparison.OrdinalIgnoreCase))
                {
                    opti = $"T.{tubeMaterial}{left}\"SQX{FormatDot(wallIn)}\"";
                }
                else if (shape.Equals("Rectangle", StringComparison.OrdinalIgnoreCase))
                {
                    opti = $"T.{tubeMaterial}{left}\"X{right}\"X{FormatDot(wallIn)}\"";
                }
                else if (shape.Equals("Angle", StringComparison.OrdinalIgnoreCase))
                {
                    opti = $"A.{tubeMaterial}{left}\"X{right}\"X{FormatDot(wallIn)}\"";
                }
                else // Channel or others
                {
                    opti = $"C.{tubeMaterial}{left}\"X{right}\"X{FormatDot(wallIn)}\"";
                }
                info.CustomProperties.SetPropertyValue("OptiMaterial", opti, CustomPropertyType.Text);
            }

            var p = _cutSvc.Get(materialCategory, wallIn);
            double cutTimeSec = (p.CutSpeedInPerMin > 0 ? (lengthIn / p.CutSpeedInPerMin) * 60.0 : 0.0);
            double cycleTimeSec = ((lengthIn / 240.0) * 45.0);
            double traverseTimeSec = (lengthIn / 1440.0) * 60.0;
            double totalHours = (cutTimeSec + cycleTimeSec + traverseTimeSec) / 3600.0;
            if (isAngleOrChannel) totalHours *= 3.0;
            info.CustomProperties.SetPropertyValue("OP20_R", totalHours.ToString("F4", CultureInfo.InvariantCulture), CustomPropertyType.Number);
            return true;
        }

        private static string ReadString(ModelInfo info, string name)
        {
            var v = info?.CustomProperties?.GetPropertyValue(name);
            return (v?.ToString() ?? string.Empty).Trim();
        }

        private static bool readDouble(ModelInfo info, string name, out double value)
        {
            value = 0.0;
            var s = ReadString(info, name);
            return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }

        private static bool TryParseDoubleInvariant(string s, out double d)
        {
            return double.TryParse((s ?? string.Empty), NumberStyles.Any, CultureInfo.InvariantCulture, out d);
        }

        private static double TryParseCrossSectionInches(string cross)
        {
            if (string.IsNullOrWhiteSpace(cross)) return 0.0;
            string cleaned = cross.Replace("\"", "").Replace("OD", "").Replace("X", " ").Replace("SCH", " ").Trim();
            if (double.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out double d))
                return d;
            foreach (var tok in cleaned.Split(new[] { ' ', '\t', ',', '*' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (double.TryParse(tok, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                    return d;
            }
            return 0.0;
        }

        private static void ParseCrossDims(string cross, out string left, out string right)
        {
            left = right = "";
            if (string.IsNullOrWhiteSpace(cross)) return;
            var raw = cross.Replace("\"", "");
            var parts = raw.Split(new[] { 'x', 'X' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 1)
            {
                if (double.TryParse(parts[0].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var val))
                {
                    left = right = FormatInchesNoSpace(val);
                }
            }
            else
            {
                if (double.TryParse(parts[0].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var l))
                {
                    left = FormatInchesNoSpace(l);
                }
                if (double.TryParse(parts[1].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var r))
                {
                    right = FormatInchesNoSpace(r);
                }
            }
        }

        private static string FormatInchesNoSpace(double v)
        {
            // Format like 2 or 1.25 or 1.315 without trailing zeros
            var s = v.ToString("0.###", CultureInfo.InvariantCulture);
            return s;
        }

        private static string FormatDot(double v)
        {
            // Leading dot for <1 values, trim trailing zeros
            string s = v.ToString("0.###", CultureInfo.InvariantCulture);
            if (s.StartsWith("0.")) s = s.Substring(1);
            return s;
        }

        private static string FormatDotText(double v)
        {
            // Return text with leading dot when <1 (e.g., .08, .2, .375)
            if (v >= 1.0) return v.ToString("0.###", CultureInfo.InvariantCulture);
            string s = v.ToString("0.###", CultureInfo.InvariantCulture);
            if (s.StartsWith("0.")) s = s.Substring(1);
            return s;
        }
    }
}
