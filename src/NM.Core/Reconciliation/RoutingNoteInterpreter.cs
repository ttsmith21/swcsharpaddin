using System.Collections.Generic;
using NM.Core.Pdf.Models;
using NM.Core.Reconciliation.Models;

namespace NM.Core.Reconciliation
{
    /// <summary>
    /// Converts drawing routing hints into sequenced routing suggestions
    /// with operation numbers matching the shop's standard routing template.
    /// </summary>
    public sealed class RoutingNoteInterpreter
    {
        // Standard operation number assignments by category.
        // These follow Northern Manufacturing's routing convention.
        private static readonly Dictionary<RoutingOp, int> DefaultOpNumbers = new Dictionary<RoutingOp, int>
        {
            { RoutingOp.ProcessOverride, 20 },  // OP20: Cutting (laser/waterjet/plasma)
            { RoutingOp.Deburr,          30 },  // OP30: Deburr
            { RoutingOp.Tap,             35 },  // OP35: Tap/drill
            { RoutingOp.Drill,           35 },  // OP35: Tap/drill
            { RoutingOp.Machine,         35 },  // OP35: Machine
            { RoutingOp.Hardware,        40 },  // OP40: Hardware insertion
            { RoutingOp.Weld,            50 },  // OP50: Welding
            { RoutingOp.Inspect,         55 },  // OP55: Inspection
            { RoutingOp.HeatTreat,       60 },  // OP60: Outside process (heat treat)
            { RoutingOp.Finish,          60 },  // OP60: Outside process (finish)
            { RoutingOp.OutsideProcess,  60 },  // OP60: Outside process (general)
        };

        /// <summary>
        /// Converts routing hints from drawing note extraction into sequenced suggestions.
        /// </summary>
        public List<RoutingSuggestion> InterpretRoutingHints(List<RoutingHint> hints)
        {
            var suggestions = new List<RoutingSuggestion>();
            if (hints == null || hints.Count == 0)
                return suggestions;

            foreach (var hint in hints)
            {
                int opNumber = DefaultOpNumbers.ContainsKey(hint.Operation)
                    ? DefaultOpNumbers[hint.Operation]
                    : 60;

                var suggestion = new RoutingSuggestion
                {
                    Operation = hint.Operation,
                    SuggestedOpNumber = opNumber,
                    WorkCenter = hint.WorkCenter,
                    NoteText = hint.NoteText ?? hint.SourceNote,
                    SourceNote = hint.SourceNote,
                    Confidence = hint.Confidence,
                    Type = DetermineType(hint)
                };

                suggestions.Add(suggestion);
            }

            // Sort by operation number for natural routing order
            suggestions.Sort((a, b) => a.SuggestedOpNumber.CompareTo(b.SuggestedOpNumber));

            return suggestions;
        }

        private static SuggestionType DetermineType(RoutingHint hint)
        {
            switch (hint.Operation)
            {
                case RoutingOp.ProcessOverride:
                    return SuggestionType.ModifyOperation;

                case RoutingOp.Deburr:
                case RoutingOp.Weld:
                case RoutingOp.Tap:
                case RoutingOp.Drill:
                case RoutingOp.Machine:
                case RoutingOp.Hardware:
                case RoutingOp.Inspect:
                    return SuggestionType.AddOperation;

                case RoutingOp.HeatTreat:
                case RoutingOp.Finish:
                case RoutingOp.OutsideProcess:
                    return string.IsNullOrEmpty(hint.WorkCenter)
                        ? SuggestionType.AddNote
                        : SuggestionType.AddOperation;

                default:
                    return SuggestionType.AddNote;
            }
        }
    }
}
