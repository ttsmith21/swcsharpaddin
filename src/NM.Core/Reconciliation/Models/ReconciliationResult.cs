using System.Collections.Generic;

namespace NM.Core.Reconciliation.Models
{
    /// <summary>
    /// Result of reconciling a 3D model (PartData) with a 2D drawing (DrawingData).
    /// Contains conflicts, gap fills, routing suggestions, and rename suggestions.
    /// </summary>
    public sealed class ReconciliationResult
    {
        /// <summary>Conflicts where 3D and 2D data disagree.</summary>
        public List<DataConflict> Conflicts { get; } = new List<DataConflict>();

        /// <summary>Fields that were empty in the model but found in the drawing.</summary>
        public List<GapFill> GapFills { get; } = new List<GapFill>();

        /// <summary>Routing operations suggested by drawing notes.</summary>
        public List<RoutingSuggestion> RoutingSuggestions { get; } = new List<RoutingSuggestion>();

        /// <summary>File rename suggestion if drawing part number differs from filename.</summary>
        public RenameSuggestion Rename { get; set; }

        /// <summary>Fields that matched between 3D and 2D (confidence boosters).</summary>
        public List<string> Confirmations { get; } = new List<string>();

        /// <summary>True if there are conflicts requiring human review.</summary>
        public bool HasConflicts => Conflicts.Count > 0;

        /// <summary>True if any data can be filled from the drawing.</summary>
        public bool HasGapFills => GapFills.Count > 0;

        /// <summary>True if drawing notes suggest routing changes.</summary>
        public bool HasRoutingSuggestions => RoutingSuggestions.Count > 0;

        /// <summary>True if a file rename is suggested.</summary>
        public bool HasRenameSuggestion => Rename != null;

        /// <summary>True if any actionable items exist.</summary>
        public bool HasActions => HasConflicts || HasGapFills || HasRoutingSuggestions || HasRenameSuggestion;

        /// <summary>Summary counts for display.</summary>
        public string Summary =>
            $"{Conflicts.Count} conflicts, {GapFills.Count} gap fills, " +
            $"{RoutingSuggestions.Count} routing suggestions, {Confirmations.Count} confirmations" +
            (HasRenameSuggestion ? ", 1 rename" : "");
    }
}
