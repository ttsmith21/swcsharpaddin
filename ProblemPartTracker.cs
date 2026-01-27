using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NM.Core
{
    /// <summary>
    /// Tracks problem parts discovered during processing. Thread-safe and NM.Core-only.
    /// </summary>
    public sealed class ProblemPartTracker
    {
        private readonly ConcurrentBag<ProblemPartInfo> _problems = new ConcurrentBag<ProblemPartInfo>();

        /// <summary>Adds a problem part using data from ModelInfo.</summary>
        public bool AddProblemPart(ModelInfo model, string reason)
        {
            const string proc = nameof(AddProblemPart);
            try
            {
                if (model == null) { ErrorHandler.HandleError(proc, "ModelInfo is null"); return false; }
                if (string.IsNullOrWhiteSpace(model.FilePath)) { ErrorHandler.HandleError(proc, "ModelInfo.FilePath is empty"); return false; }
                if (string.IsNullOrWhiteSpace(reason)) { ErrorHandler.HandleError(proc, "Reason is empty"); return false; }

                var cfg = model.ConfigurationName ?? string.Empty;
                var info = new ProblemPartInfo(model.FilePath, reason, cfg);
                _problems.Add(info);
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Exception adding problem part", ex);
                return false;
            }
        }

        /// <summary>Removes a matching problem entry using path+configuration equality.</summary>
        public bool RemoveProblemPart(ProblemPartInfo item)
        {
            const string proc = nameof(RemoveProblemPart);
            try
            {
                if (item == null) return false;
                // ConcurrentBag doesn't support removal; rebuild without the item.
                var remaining = _problems.Where(p => !p.Equals(item)).ToList();
                // Swap contents
                while (_problems.TryTake(out _)) { }
                foreach (var r in remaining) _problems.Add(r);
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.HandleError(proc, "Exception removing problem part", ex);
                return false;
            }
        }

        /// <summary>Returns a snapshot of all problems.</summary>
        public IReadOnlyList<ProblemPartInfo> GetAllProblems()
        {
            return _problems.ToArray();
        }

        /// <summary>Clears all recorded problems.</summary>
        public void ClearAll()
        {
            while (_problems.TryTake(out _)) { }
        }

        /// <summary>Returns a multi-line summary suitable for UI display.</summary>
        public string GetDisplaySummary()
        {
            var list = GetAllProblems();
            if (list.Count == 0) return "No problems found.";
            var sb = new StringBuilder(list.Count * 64);
            foreach (var p in list.OrderBy(p => p.PartPath, StringComparer.OrdinalIgnoreCase)
                                   .ThenBy(p => p.Configuration, StringComparer.OrdinalIgnoreCase))
            {
                sb.AppendLine(p.GetDisplayText());
            }
            return sb.ToString();
        }
    }
}
