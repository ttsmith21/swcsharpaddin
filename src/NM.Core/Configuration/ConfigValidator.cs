using System.Collections.Generic;

namespace NM.Core.Config
{
    /// <summary>
    /// Validates loaded configuration values are within reasonable ranges.
    /// Called by NmConfigProvider after deserialization.
    /// </summary>
    public static class ConfigValidator
    {
        public sealed class ValidationMessage
        {
            public string Field { get; set; }
            public string Message { get; set; }
            public bool IsError { get; set; }

            public override string ToString() => $"[{(IsError ? "ERROR" : "WARN")}] {Field}: {Message}";
        }

        public static List<ValidationMessage> Validate(NmConfig config)
        {
            var msgs = new List<ValidationMessage>();
            if (config == null)
            {
                msgs.Add(Error("NmConfig", "Config object is null"));
                return msgs;
            }

            // Work center rates: must be > 0 and < 1000 $/hr
            ValidateRate(msgs, "WorkCenters.F115", config.WorkCenters.F115_LaserCutting);
            ValidateRate(msgs, "WorkCenters.F140", config.WorkCenters.F140_PressBrake);
            ValidateRate(msgs, "WorkCenters.F145", config.WorkCenters.F145_CncBending);
            ValidateRate(msgs, "WorkCenters.F210", config.WorkCenters.F210_Deburring);
            ValidateRate(msgs, "WorkCenters.F220", config.WorkCenters.F220_Tapping);
            ValidateRate(msgs, "WorkCenters.F300", config.WorkCenters.F300_MaterialHandling);
            ValidateRate(msgs, "WorkCenters.F325", config.WorkCenters.F325_RollForming);
            ValidateRate(msgs, "WorkCenters.F385", config.WorkCenters.F385_Assembly);
            ValidateRate(msgs, "WorkCenters.F400", config.WorkCenters.F400_Welding);
            ValidateRate(msgs, "WorkCenters.F500", config.WorkCenters.F500_Finishing);
            ValidateRate(msgs, "WorkCenters.F525", config.WorkCenters.F525_Packaging);
            ValidateRate(msgs, "WorkCenters.F155", config.WorkCenters.F155_Waterjet);
            ValidateRate(msgs, "WorkCenters.ENG", config.WorkCenters.ENG_Engineering);

            // Material pricing: must be > 0 and < 100 $/lb
            ValidatePrice(msgs, "MaterialPricing.Stainless304", config.MaterialPricing.Stainless304_PerLb);
            ValidatePrice(msgs, "MaterialPricing.Stainless316", config.MaterialPricing.Stainless316_PerLb);
            ValidatePrice(msgs, "MaterialPricing.CarbonSteel", config.MaterialPricing.CarbonSteel_PerLb);
            ValidatePrice(msgs, "MaterialPricing.Aluminum6061", config.MaterialPricing.Aluminum6061_PerLb);
            ValidatePrice(msgs, "MaterialPricing.Aluminum5052", config.MaterialPricing.Aluminum5052_PerLb);
            ValidatePrice(msgs, "MaterialPricing.Galvanized", config.MaterialPricing.Galvanized_PerLb);

            // Pricing modifiers: markup should be 1.0-2.0
            if (config.Pricing.MaterialMarkup < 1.0 || config.Pricing.MaterialMarkup > 2.0)
                msgs.Add(Warn("Pricing.MaterialMarkup", $"Value {config.Pricing.MaterialMarkup} outside expected range 1.0-2.0"));

            // Press brake rate arrays should have 5 and 3 entries respectively
            if (config.Manufacturing.PressBrake.RateSeconds == null || config.Manufacturing.PressBrake.RateSeconds.Length != 5)
                msgs.Add(Error("Manufacturing.PressBrake.RateSeconds", "Expected exactly 5 rate entries"));
            if (config.Manufacturing.PressBrake.WeightThresholdsLbs == null || config.Manufacturing.PressBrake.WeightThresholdsLbs.Length != 3)
                msgs.Add(Error("Manufacturing.PressBrake.WeightThresholdsLbs", "Expected exactly 3 weight thresholds"));
            if (config.Manufacturing.PressBrake.LengthThresholdsIn == null || config.Manufacturing.PressBrake.LengthThresholdsIn.Length != 2)
                msgs.Add(Error("Manufacturing.PressBrake.LengthThresholdsIn", "Expected exactly 2 length thresholds"));

            // Densities must be > 0
            if (config.MaterialDensities.Stainless304_316 <= 0)
                msgs.Add(Error("MaterialDensities.Stainless304_316", "Density must be > 0"));
            if (config.MaterialDensities.CarbonSteelA36 <= 0)
                msgs.Add(Error("MaterialDensities.CarbonSteelA36", "Density must be > 0"));
            if (config.MaterialDensities.Aluminum <= 0)
                msgs.Add(Error("MaterialDensities.Aluminum", "Density must be > 0"));
            if (config.MaterialDensities.Steel_General <= 0)
                msgs.Add(Error("MaterialDensities.Steel_General", "Density must be > 0"));

            // Standard sheet dimensions
            if (config.Manufacturing.StandardSheet.WidthIn <= 0 || config.Manufacturing.StandardSheet.LengthIn <= 0)
                msgs.Add(Error("Manufacturing.StandardSheet", "Sheet dimensions must be > 0"));

            // Schema version
            if (config.SchemaVersion < 1)
                msgs.Add(Warn("SchemaVersion", $"Unexpected schema version {config.SchemaVersion}"));

            return msgs;
        }

        private static void ValidateRate(List<ValidationMessage> msgs, string field, double value)
        {
            if (value <= 0)
                msgs.Add(Error(field, $"Rate must be > 0, got {value}"));
            else if (value > 1000)
                msgs.Add(Warn(field, $"Rate {value} $/hr seems high (> $1000/hr)"));
        }

        private static void ValidatePrice(List<ValidationMessage> msgs, string field, double value)
        {
            if (value <= 0)
                msgs.Add(Error(field, $"Price must be > 0, got {value}"));
            else if (value > 100)
                msgs.Add(Warn(field, $"Price {value} $/lb seems high (> $100/lb)"));
        }

        private static ValidationMessage Error(string field, string message)
            => new ValidationMessage { Field = field, Message = message, IsError = true };

        private static ValidationMessage Warn(string field, string message)
            => new ValidationMessage { Field = field, Message = message, IsError = false };
    }
}
