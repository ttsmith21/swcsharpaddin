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
    /// For parts: maps to standard property names (OP20_S, F210_S, F220_Note, etc.).
    /// For assemblies: generates free-form OP## operation slots.
    /// These suggestions are presented in the UI wizard for user approval
    /// before being written to the model via CustomPropertiesService.
    /// </summary>
    public sealed class PropertySuggestionService
    {
        /// <summary>
        /// Generates property suggestions for a PART from reconciliation results.
        /// Maps to the standard part custom property schema (OP20, F210, F140, F220, F325).
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
                string propName = MapFieldToProperty(gap.Field);
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
            foreach (var suggestion in reconciliation.RoutingSuggestions)
            {
                var routingProps = MapRoutingSuggestionToPartProperties(suggestion, current);
                suggestions.AddRange(routingProps);
            }

            // 3. Conflict resolutions (where we have a recommendation)
            foreach (var conflict in reconciliation.Conflicts.Where(c =>
                c.Recommendation != ConflictResolution.HumanRequired))
            {
                string propName = MapFieldToProperty(conflict.Field);
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

            // 1. Identity fields (same as parts)
            foreach (var gap in reconciliation.GapFills)
            {
                string propName = MapFieldToProperty(gap.Field);
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

        // --- Mapping helpers ---

        /// <summary>
        /// Maps a routing suggestion to the correct part custom property names.
        /// Deburr → F210_*, Weld → F400_*, Tap → F220_*, etc.
        /// </summary>
        private List<PropertySuggestion> MapRoutingSuggestionToPartProperties(
            RoutingSuggestion suggestion, IDictionary<string, string> current)
        {
            var props = new List<PropertySuggestion>();

            switch (suggestion.Operation)
            {
                case RoutingOp.Deburr:
                    // F210 deburr — add a routing note to existing deburr operation
                    props.Add(MakeRoutingNote("F210_Note", suggestion, current));
                    break;

                case RoutingOp.Tap:
                case RoutingOp.Drill:
                    // F220 tapping — add note to existing tap operation
                    props.Add(MakeRoutingNote("F220_Note", suggestion, current));
                    break;

                case RoutingOp.Weld:
                    // Welding — not a standard part operation, add to Extra
                    props.Add(MakeRoutingNote("WeldNote", suggestion, current));
                    if (!string.IsNullOrEmpty(suggestion.WorkCenter))
                        props.Add(MakeProperty("WeldWorkCenter", suggestion.WorkCenter,
                            current, SuggestionSource.DrawingNote, suggestion.Confidence,
                            $"Weld operation from drawing note: \"{suggestion.SourceNote}\""));
                    break;

                case RoutingOp.ProcessOverride:
                    // Override the OP20 work center
                    if (!string.IsNullOrEmpty(suggestion.WorkCenter))
                    {
                        props.Add(MakeProperty("OP20_WorkCenter", suggestion.WorkCenter,
                            current, SuggestionSource.DrawingNote, suggestion.Confidence,
                            $"Process override from drawing: \"{suggestion.SourceNote}\""));
                    }
                    break;

                case RoutingOp.Finish:
                case RoutingOp.HeatTreat:
                case RoutingOp.OutsideProcess:
                    // Outside processing — routing note for ERP
                    props.Add(MakeRoutingNote("OutsideProcessNote", suggestion, current));
                    break;

                case RoutingOp.Inspect:
                    props.Add(MakeRoutingNote("InspectionNote", suggestion, current));
                    break;

                case RoutingOp.Hardware:
                    props.Add(MakeRoutingNote("HardwareNote", suggestion, current));
                    break;

                case RoutingOp.Machine:
                    props.Add(MakeRoutingNote("MachineNote", suggestion, current));
                    break;
            }

            return props;
        }

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

        private static string MapFieldToProperty(string field)
        {
            switch (field)
            {
                case "PartNumber": return "PartNumber";
                case "Description": return "Description";
                case "Revision": return "Revision";
                case "Material": return "rbMaterialType";
                case "Finish": return "Finish";
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
            // Default setup estimates in minutes (user can adjust in UI)
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
            // Default run estimates in minutes (user can adjust in UI)
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
