using System;
using System.Collections.Generic;
using System.Linq;
using NM.Core.DataModel;
using NM.Core.Pdf.Models;
using NM.Core.Reconciliation.Models;

namespace NM.Core.Reconciliation
{
    /// <summary>
    /// Generates SolidWorks custom property suggestions from reconciliation results.
    /// Property names match the exact SolidWorks custom property tab builder schema.
    ///
    /// PARTS fixed slots: OP20 (cut), F210/OP30 (deburr), F220/OP35 (tap),
    ///   F140/OP40 (brake), F325/OP50 (roll), plus 6 "Other WC" slots and outsource.
    /// ASSEMBLIES free-form: OP20..OP150 (each with WC, setup, run, note).
    ///
    /// See docs/reference/part-tab-builder-schema.md and assembly-tab-builder-schema.md.
    /// </summary>
    public sealed class PropertySuggestionService
    {
        /// <summary>
        /// Generates property suggestions for a PART from reconciliation results.
        /// Maps to the standard part custom property schema.
        /// </summary>
        public List<PropertySuggestion> GeneratePartSuggestions(
            ReconciliationResult reconciliation,
            IDictionary<string, string> currentProperties)
        {
            var suggestions = new List<PropertySuggestion>();
            var current = currentProperties ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 1. Identity fields from gap fills
            foreach (var gap in reconciliation.GapFills)
            {
                string propName = MapFieldToPartProperty(gap.Field);
                if (propName == null) continue;

                suggestions.Add(new PropertySuggestion
                {
                    PropertyName = propName,
                    Value = gap.Value,
                    CurrentValue = GetCurrent(current, propName),
                    Source = SuggestionSource.DrawingTitleBlock,
                    Category = GetCategory(gap.Field),
                    Confidence = gap.Confidence,
                    Reason = $"Extracted from PDF drawing ({gap.Source})"
                });
            }

            // 2. Routing notes from drawing notes → specific operation properties
            // Track "Other WC" slot allocation across all suggestions
            int nextOtherSlot = 1;

            foreach (var suggestion in reconciliation.RoutingSuggestions)
            {
                var routingProps = MapRoutingSuggestionToPartProperties(
                    suggestion, current, ref nextOtherSlot);
                suggestions.AddRange(routingProps);
            }

            // 3. Conflict resolutions (where we have a recommendation)
            foreach (var conflict in reconciliation.Conflicts.Where(c =>
                c.Recommendation != ConflictResolution.HumanRequired))
            {
                string propName = MapFieldToPartProperty(conflict.Field);
                if (propName == null) continue;

                string value = conflict.Recommendation == ConflictResolution.UseDrawing
                    ? conflict.DrawingValue
                    : conflict.ModelValue;

                suggestions.Add(new PropertySuggestion
                {
                    PropertyName = propName,
                    Value = value,
                    CurrentValue = conflict.ModelValue,
                    Source = SuggestionSource.Reconciliation,
                    Category = GetCategory(conflict.Field),
                    Confidence = 0.7,
                    Reason = conflict.Reason
                });
            }

            return suggestions;
        }

        /// <summary>
        /// Generates operation slot suggestions for an ASSEMBLY from drawing notes.
        /// Assemblies use free-form OP## slots (OP10 = KIT, then OP20+).
        /// Work center is stored in OP## (ComboBox), not OP##_WC.
        /// </summary>
        public List<AssemblyOperationSuggestion> GenerateAssemblyOperations(
            ReconciliationResult reconciliation,
            int startingOpNumber = 20)
        {
            var operations = new List<AssemblyOperationSuggestion>();

            if (!reconciliation.HasRoutingSuggestions)
                return operations;

            int nextOp = startingOpNumber;

            foreach (var suggestion in reconciliation.RoutingSuggestions)
            {
                operations.Add(new AssemblyOperationSuggestion
                {
                    OpNumber = nextOp,
                    WorkCenter = suggestion.WorkCenter ?? MapOperationToDefaultWorkCenter(suggestion.Operation),
                    Setup_min = EstimateSetup(suggestion.Operation),
                    Run_min = EstimateRun(suggestion.Operation),
                    RoutingNote = suggestion.NoteText,
                    SourceNote = suggestion.SourceNote,
                    Confidence = suggestion.Confidence
                });

                nextOp += 10;
            }

            return operations;
        }

        /// <summary>
        /// Generates the full set of property suggestions for an assembly,
        /// including identity fields and operation slots.
        /// </summary>
        public List<PropertySuggestion> GenerateAssemblySuggestions(
            ReconciliationResult reconciliation,
            IDictionary<string, string> currentProperties,
            int startingOpNumber = 20)
        {
            var suggestions = new List<PropertySuggestion>();
            var current = currentProperties ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 1. Identity fields (same mapping as parts)
            foreach (var gap in reconciliation.GapFills)
            {
                string propName = MapFieldToPartProperty(gap.Field);
                if (propName == null) continue;

                suggestions.Add(new PropertySuggestion
                {
                    PropertyName = propName,
                    Value = gap.Value,
                    CurrentValue = GetCurrent(current, propName),
                    Source = SuggestionSource.DrawingTitleBlock,
                    Category = GetCategory(gap.Field),
                    Confidence = gap.Confidence,
                    Reason = $"Extracted from PDF drawing ({gap.Source})"
                });
            }

            // 2. Operation slots → flat property suggestions
            var operations = GenerateAssemblyOperations(reconciliation, startingOpNumber);
            foreach (var op in operations)
            {
                var props = op.ToProperties();
                foreach (var kv in props)
                {
                    suggestions.Add(new PropertySuggestion
                    {
                        PropertyName = kv.Key,
                        Value = kv.Value,
                        CurrentValue = GetCurrent(current, kv.Key),
                        Source = SuggestionSource.DrawingNote,
                        Category = PropertyCategory.Routing,
                        Confidence = op.Confidence,
                        Reason = $"Drawing note: \"{op.SourceNote}\""
                    });
                }
            }

            return suggestions;
        }

        // --- Part routing mapping ---

        /// <summary>
        /// Maps a routing suggestion to the correct part custom property names.
        /// Fixed slots: Deburr→F210_RN, Tap→F220_RN, ProcessOverride→OP20.
        /// Flexible: Weld/Hardware/Machine/Inspect→Other WC slots.
        /// Outside: Finish/HeatTreat/OutsideProcess→OS_WC/OS_OP/OS_RN.
        /// </summary>
        private List<PropertySuggestion> MapRoutingSuggestionToPartProperties(
            RoutingSuggestion suggestion, IDictionary<string, string> current,
            ref int nextOtherSlot)
        {
            var props = new List<PropertySuggestion>();

            switch (suggestion.Operation)
            {
                case RoutingOp.Deburr:
                    // F210 deburr (OP30) — enable checkbox + routing note
                    props.Add(MakeProperty("F210", "1", current,
                        SuggestionSource.DrawingNote, suggestion.Confidence,
                        $"Enable deburr from drawing note: \"{suggestion.SourceNote}\""));
                    props.Add(MakeRoutingNote("F210_RN", suggestion, current));
                    break;

                case RoutingOp.Tap:
                case RoutingOp.Drill:
                    // F220 tap/drill (OP35) — enable checkbox + routing note
                    props.Add(MakeProperty("F220", "1", current,
                        SuggestionSource.DrawingNote, suggestion.Confidence,
                        $"Enable tap/drill from drawing note: \"{suggestion.SourceNote}\""));
                    props.Add(MakeRoutingNote("F220_RN", suggestion, current));
                    break;

                case RoutingOp.ProcessOverride:
                    // Override the OP20 work center (ComboBox property is "OP20")
                    if (!string.IsNullOrEmpty(suggestion.WorkCenter))
                    {
                        props.Add(MakeProperty("OP20", suggestion.WorkCenter,
                            current, SuggestionSource.DrawingNote, suggestion.Confidence,
                            $"Process override from drawing: \"{suggestion.SourceNote}\""));
                    }
                    if (!string.IsNullOrEmpty(suggestion.NoteText))
                    {
                        props.Add(MakeRoutingNote("OP20_RN", suggestion, current));
                    }
                    break;

                case RoutingOp.Finish:
                case RoutingOp.HeatTreat:
                case RoutingOp.OutsideProcess:
                    // Outside processing → OS_WC, OS_OP, OS_RN
                    if (!string.IsNullOrEmpty(suggestion.WorkCenter))
                    {
                        props.Add(MakeProperty("OS_WC", suggestion.WorkCenter,
                            current, SuggestionSource.DrawingNote, suggestion.Confidence,
                            $"Outside process from drawing: \"{suggestion.SourceNote}\""));
                    }
                    props.Add(MakeRoutingNote("OS_RN", suggestion, current));
                    break;

                case RoutingOp.Weld:
                case RoutingOp.Inspect:
                case RoutingOp.Hardware:
                case RoutingOp.Machine:
                    // These use "Other WC" slots (1-6) on parts
                    if (nextOtherSlot <= 6)
                    {
                        var slotProps = MakeOtherWcSlot(
                            nextOtherSlot, suggestion, current);
                        props.AddRange(slotProps);
                        nextOtherSlot++;
                    }
                    break;
            }

            return props;
        }

        /// <summary>
        /// Creates property suggestions for an "Other WC" slot (1-6).
        /// Slot 1: OtherWC_CB, OtherOP, Other_WC, Other_S, Other_R, Other_RN
        /// Slot N (2-6): OtherWC_CB{N}, Other_OP{N}, Other_WC{N}, Other_S{N}, Other_R{N}, Other_RN{N}
        /// </summary>
        private List<PropertySuggestion> MakeOtherWcSlot(
            int slotNumber, RoutingSuggestion suggestion, IDictionary<string, string> current)
        {
            var props = new List<PropertySuggestion>();
            string reason = $"Drawing note: \"{suggestion.SourceNote}\"";

            // Property name suffixes differ between slot 1 and 2-6
            string cbName = slotNumber == 1 ? "OtherWC_CB" : $"OtherWC_CB{slotNumber}";
            string opName = slotNumber == 1 ? "OtherOP" : $"Other_OP{slotNumber}";
            string wcName = slotNumber == 1 ? "Other_WC" : $"Other_WC{slotNumber}";
            string sName = slotNumber == 1 ? "Other_S" : $"Other_S{slotNumber}";
            string rName = slotNumber == 1 ? "Other_R" : $"Other_R{slotNumber}";
            string rnName = slotNumber == 1 ? "Other_RN" : $"Other_RN{slotNumber}";

            // Enable the checkbox
            props.Add(MakeProperty(cbName, "1", current,
                SuggestionSource.DrawingNote, suggestion.Confidence, reason));

            // Set the operation number (default: 60 + (slot-1)*10)
            int defaultOp = 60 + (slotNumber - 1) * 10;
            props.Add(MakeProperty(opName, defaultOp.ToString(),
                current, SuggestionSource.DrawingNote, suggestion.Confidence, reason));

            // Work center
            string wc = suggestion.WorkCenter ?? MapOperationToDefaultWorkCenter(suggestion.Operation);
            if (!string.IsNullOrEmpty(wc))
            {
                props.Add(MakeProperty(wcName, wc,
                    current, SuggestionSource.DrawingNote, suggestion.Confidence, reason));
            }

            // Setup and run estimates
            double setup = EstimateSetup(suggestion.Operation);
            double run = EstimateRun(suggestion.Operation);
            if (setup > 0)
                props.Add(MakeProperty(sName, setup.ToString("0.####",
                    System.Globalization.CultureInfo.InvariantCulture),
                    current, SuggestionSource.DrawingNote, suggestion.Confidence, reason));
            if (run > 0)
                props.Add(MakeProperty(rName, run.ToString("0.####",
                    System.Globalization.CultureInfo.InvariantCulture),
                    current, SuggestionSource.DrawingNote, suggestion.Confidence, reason));

            // Routing note
            props.Add(MakeRoutingNote(rnName, suggestion, current));

            return props;
        }

        // --- Shared helpers ---

        private PropertySuggestion MakeRoutingNote(
            string propertyName, RoutingSuggestion suggestion, IDictionary<string, string> current)
        {
            string existingNote = GetCurrent(current, propertyName);
            string newValue = string.IsNullOrEmpty(existingNote)
                ? suggestion.NoteText
                : existingNote + "; " + suggestion.NoteText;

            return new PropertySuggestion
            {
                PropertyName = propertyName,
                Value = newValue,
                CurrentValue = existingNote,
                Source = SuggestionSource.DrawingNote,
                Category = PropertyCategory.Routing,
                Confidence = suggestion.Confidence,
                Reason = $"Drawing note: \"{suggestion.SourceNote}\""
            };
        }

        private PropertySuggestion MakeProperty(
            string propertyName, string value, IDictionary<string, string> current,
            SuggestionSource source, double confidence, string reason)
        {
            return new PropertySuggestion
            {
                PropertyName = propertyName,
                Value = value,
                CurrentValue = GetCurrent(current, propertyName),
                Source = source,
                Category = PropertyCategory.Routing,
                Confidence = confidence,
                Reason = reason
            };
        }

        // --- Field → property name mapping ---

        /// <summary>
        /// Maps reconciliation field names to SolidWorks custom property names.
        /// Based on the part/assembly tab builder schema.
        /// </summary>
        private static string MapFieldToPartProperty(string field)
        {
            switch (field)
            {
                case "PartNumber": return "Print";       // Tab builder uses "Print" for part number
                case "Description": return "Description";
                case "Revision": return "Revision";
                case "Material": return "OptiMaterial";  // Material name → OptiMaterial (NOT rbMaterialType which is a RadioButton 0/1/2)
                case "Finish": return "OS_RN";           // Finish notes go to outside service routing note
                default: return null;
            }
        }

        private static PropertyCategory GetCategory(string field)
        {
            switch (field)
            {
                case "PartNumber":
                case "Description":
                case "Revision":
                    return PropertyCategory.Identity;
                case "Material":
                case "Finish":
                    return PropertyCategory.Material;
                default:
                    return PropertyCategory.Routing;
            }
        }

        private static string GetCurrent(IDictionary<string, string> current, string key)
        {
            if (current != null && current.TryGetValue(key, out string value))
                return value;
            return null;
        }

        private static string MapOperationToDefaultWorkCenter(RoutingOp op)
        {
            switch (op)
            {
                case RoutingOp.Deburr: return "F210";
                case RoutingOp.Weld: return "F400";
                case RoutingOp.Tap:
                case RoutingOp.Drill: return "F220";
                case RoutingOp.Machine: return "F220";
                case RoutingOp.Hardware: return "F220";
                default: return null;
            }
        }

        private static double EstimateSetup(RoutingOp op)
        {
            switch (op)
            {
                case RoutingOp.Weld: return 15;
                case RoutingOp.Deburr: return 5;
                case RoutingOp.Tap:
                case RoutingOp.Drill: return 10;
                case RoutingOp.Machine: return 15;
                case RoutingOp.Inspect: return 5;
                default: return 0;
            }
        }

        private static double EstimateRun(RoutingOp op)
        {
            switch (op)
            {
                case RoutingOp.Weld: return 10;
                case RoutingOp.Deburr: return 2;
                case RoutingOp.Tap:
                case RoutingOp.Drill: return 5;
                case RoutingOp.Machine: return 10;
                case RoutingOp.Inspect: return 5;
                default: return 0;
            }
        }
    }
}
