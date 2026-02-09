using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using NM.Core.Pdf.Models;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Extracts and classifies manufacturing notes from engineering drawing text.
    /// Converts raw text into categorized DrawingNote objects and RoutingHints.
    /// </summary>
    public sealed class DrawingNoteExtractor
    {
        private static readonly NotePattern[] Patterns = BuildPatterns();

        private static readonly Regex BomHeaderRegex = new Regex(
            @"\b(?:BILL\s*OF\s*MATERIAL|BOM|PARTS?\s*LIST|ITEM\s+(?:NO|#|NUMBER))\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex BomRowRegex = new Regex(
            @"^\s*\d{1,3}\s+[A-Z0-9][\w\-\.]+\s+.+?\s+\d{1,4}\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.Multiline);

        /// <summary>
        /// Replaces BOM table row text with spaces to prevent false positives,
        /// while preserving character positions for the overlap-aware NoteCandidate system.
        /// </summary>
        private static string StripBomTableText(string text)
        {
            if (!BomHeaderRegex.IsMatch(text)) return text;
            return BomRowRegex.Replace(text, m => new string(' ', m.Length));
        }

        /// <summary>
        /// Extracts manufacturing-relevant notes from page text.
        /// </summary>
        public List<DrawingNote> ExtractNotes(string text, int pageNumber = 1)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new List<DrawingNote>();

            // Strip BOM table rows to prevent part descriptions from matching as notes
            string strippedText = StripBomTableText(text);

            var notes = new List<DrawingNote>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            // Collect all candidate matches with their positions, then deduplicate
            var candidates = new List<NoteCandidate>();

            foreach (var pattern in Patterns)
            {
                var matches = pattern.Regex.Matches(strippedText);
                foreach (Match match in matches)
                {
                    candidates.Add(new NoteCandidate
                    {
                        Start = match.Index,
                        End = match.Index + match.Length,
                        Text = match.Value.Trim(),
                        Pattern = pattern,
                        PageNumber = pageNumber
                    });
                }
            }

            // Sort by length descending so longer (more specific) matches take priority
            candidates.Sort((a, b) => (b.End - b.Start).CompareTo(a.End - a.Start));

            var matchedRanges = new List<int[]>();
            foreach (var c in candidates)
            {
                if (seen.Contains(c.Text)) continue;

                // Skip if this match overlaps with an already-accepted range
                bool overlaps = false;
                foreach (var range in matchedRanges)
                {
                    if (c.Start < range[1] && c.End > range[0])
                    {
                        overlaps = true;
                        break;
                    }
                }
                if (overlaps) continue;

                seen.Add(c.Text);
                matchedRanges.Add(new[] { c.Start, c.End });

                notes.Add(new DrawingNote
                {
                    Text = c.Text,
                    Category = c.Pattern.Category,
                    Impact = c.Pattern.Impact,
                    Confidence = c.Pattern.Confidence,
                    PageNumber = c.PageNumber
                });
            }

            // Also extract numbered notes (common format: "1. BREAK ALL EDGES", "NOTE: ...")
            ExtractNumberedNotes(text, pageNumber, notes, seen);

            return notes;
        }

        /// <summary>
        /// Converts extracted notes into routing hints with work center assignments.
        /// </summary>
        public List<RoutingHint> GenerateRoutingHints(List<DrawingNote> notes)
        {
            if (notes == null || notes.Count == 0)
                return new List<RoutingHint>();

            var hints = new List<RoutingHint>();

            foreach (var note in notes)
            {
                if (note.Impact == RoutingImpact.Informational)
                    continue;

                foreach (var pattern in Patterns)
                {
                    if (!pattern.Regex.IsMatch(note.Text))
                        continue;

                    if (string.IsNullOrEmpty(pattern.WorkCenter) && pattern.RoutingOp == RoutingOp.OutsideProcess)
                    {
                        // Outside process — generate note text from pattern
                        string routingNote = pattern.RoutingNoteTemplate;
                        if (!string.IsNullOrEmpty(routingNote))
                        {
                            var match = pattern.Regex.Match(note.Text);
                            routingNote = SubstituteGroups(routingNote, match);
                        }
                        else
                        {
                            routingNote = note.Text.ToUpperInvariant();
                        }

                        hints.Add(new RoutingHint
                        {
                            Operation = pattern.RoutingOp,
                            WorkCenter = null,
                            NoteText = routingNote,
                            SourceNote = note.Text,
                            Confidence = note.Confidence
                        });
                    }
                    else
                    {
                        hints.Add(new RoutingHint
                        {
                            Operation = pattern.RoutingOp,
                            WorkCenter = pattern.WorkCenter,
                            NoteText = pattern.RoutingNoteTemplate ?? note.Text.ToUpperInvariant(),
                            SourceNote = note.Text,
                            Confidence = note.Confidence
                        });
                    }

                    break; // First matching pattern wins
                }
            }

            return hints;
        }

        private void ExtractNumberedNotes(string text, int pageNumber, List<DrawingNote> notes, HashSet<string> seen)
        {
            // Match "NOTES:" section followed by numbered items
            var notesSection = Regex.Match(text, @"NOTES?\s*:\s*\r?\n((?:\s*\d+[\.\)]\s*.+\r?\n?)+)", RegexOptions.IgnoreCase);
            if (!notesSection.Success) return;

            var numberedNotes = Regex.Matches(notesSection.Groups[1].Value, @"\d+[\.\)]\s*(.+?)(?=\r?\n\s*\d+[\.\)]|\s*$)", RegexOptions.IgnoreCase);
            foreach (Match match in numberedNotes)
            {
                string noteText = match.Groups[1].Value.Trim();
                if (string.IsNullOrWhiteSpace(noteText) || seen.Contains(noteText)) continue;
                seen.Add(noteText);

                // Classify the numbered note
                var category = ClassifyNote(noteText);
                notes.Add(new DrawingNote
                {
                    Text = noteText,
                    Category = category,
                    Impact = category == NoteCategory.General ? RoutingImpact.Informational : RoutingImpact.AddOperation,
                    Confidence = 0.75,
                    PageNumber = pageNumber
                });
            }
        }

        private NoteCategory ClassifyNote(string text)
        {
            foreach (var pattern in Patterns)
            {
                if (pattern.Regex.IsMatch(text))
                    return pattern.Category;
            }
            return NoteCategory.General;
        }

        private static string SubstituteGroups(string template, Match match)
        {
            if (match == null || !match.Success) return template;
            for (int i = 0; i < match.Groups.Count; i++)
            {
                template = template.Replace("{" + i + "}", match.Groups[i].Value.Trim().ToUpperInvariant());
            }
            return template;
        }

        private static NotePattern[] BuildPatterns()
        {
            return new[]
            {
                // DEBURR / EDGE BREAK
                new NotePattern(@"\bbreak\s+(all\s+)?(?:sharp\s+)?edges\b", NoteCategory.Deburr, RoutingImpact.AddOperation, RoutingOp.Deburr, "F210", "BREAK ALL EDGES", 0.95),
                new NotePattern(@"\bdeburr\b\s*(all)?", NoteCategory.Deburr, RoutingImpact.AddOperation, RoutingOp.Deburr, "F210", "DEBURR", 0.95),
                new NotePattern(@"\bremove\s+(all\s+)?burrs\b", NoteCategory.Deburr, RoutingImpact.AddOperation, RoutingOp.Deburr, "F210", "REMOVE ALL BURRS", 0.95),
                new NotePattern(@"\btumble\s*deburr\b", NoteCategory.Deburr, RoutingImpact.AddOperation, RoutingOp.Deburr, "F210", "TUMBLE DEBURR", 0.90),
                new NotePattern(@"\bradius\s+all\s+edges\b", NoteCategory.Deburr, RoutingImpact.AddOperation, RoutingOp.Deburr, "F210", "RADIUS ALL EDGES", 0.90),

                // SURFACE FINISH / COATING
                new NotePattern(@"\bpaint\s+(?!can\b|booth\b|gun\b|shop\b|room\b|line\b|mask\b)(.{3,60}?)(?:\s*$|\s*per\s)", NoteCategory.Finish, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "PAINT {1}", 0.90),
                new NotePattern(@"\bpowder\s*coat\s*(.{0,60}?)(?:\s*$)", NoteCategory.Finish, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "POWDER COAT", 0.90),
                new NotePattern(@"\banodize\s*(.{0,60}?)(?:\s*$)", NoteCategory.Finish, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "ANODIZE", 0.90),
                new NotePattern(@"\bgalvanize\b(?!d\s+(?:steel|iron|metal|sheet|angle|channel))", NoteCategory.Finish, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "GALVANIZE", 0.90),
                new NotePattern(@"\bzinc\s*plate\b", NoteCategory.Finish, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "ZINC PLATE", 0.90),
                new NotePattern(@"\bchrome\s*plate\b", NoteCategory.Finish, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "CHROME PLATE", 0.90),
                new NotePattern(@"\bblack\s*oxide\b", NoteCategory.Finish, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "BLACK OXIDE", 0.90),
                new NotePattern(@"\bhot\s*dip\s*galv", NoteCategory.Finish, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "HOT DIP GALVANIZE", 0.90),
                new NotePattern(@"\be-?coat\b", NoteCategory.Finish, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "E-COAT", 0.85),
                new NotePattern(@"\bprime[rd]?\s+(?:coat|all|surfaces?|before|and|per|with)\b", NoteCategory.Finish, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "PRIME", 0.80),

                // HEAT TREAT
                new NotePattern(@"\bheat\s*treat\s*(.{0,60}?)(?:\s*$)", NoteCategory.HeatTreat, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "HEAT TREAT", 0.90),
                new NotePattern(@"\bstress\s*reliev", NoteCategory.HeatTreat, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "STRESS RELIEVE", 0.90),
                new NotePattern(@"\bharden\s*(?:to|per)\b", NoteCategory.HeatTreat, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "HARDEN", 0.90),
                new NotePattern(@"\b(?:RC|HRC|ROCKWELL)\s*(\d{2})", NoteCategory.HeatTreat, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "HARDEN TO {0}", 0.85),
                new NotePattern(@"\bnormalize\b(?!\s+data)", NoteCategory.HeatTreat, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "NORMALIZE", 0.85),
                new NotePattern(@"\banneal\b", NoteCategory.HeatTreat, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "ANNEAL", 0.85),
                new NotePattern(@"\bcarburize\b", NoteCategory.HeatTreat, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "CARBURIZE", 0.85),
                new NotePattern(@"\bcase\s*harden\b", NoteCategory.HeatTreat, RoutingImpact.AddOperation, RoutingOp.OutsideProcess, null, "CASE HARDEN", 0.85),

                // WELDING
                new NotePattern(@"\bweld\s*(?:all|per|as)\b", NoteCategory.Weld, RoutingImpact.AddOperation, RoutingOp.Weld, "F400", "WELD PER DWG", 0.90),
                new NotePattern(@"\bmig\s*weld\b", NoteCategory.Weld, RoutingImpact.AddOperation, RoutingOp.Weld, "F400", "MIG WELD", 0.90),
                new NotePattern(@"\btig\s*weld\b", NoteCategory.Weld, RoutingImpact.AddOperation, RoutingOp.Weld, "F400", "TIG WELD", 0.90),
                new NotePattern(@"\bspot\s*weld\b", NoteCategory.Weld, RoutingImpact.AddOperation, RoutingOp.Weld, "F400", "SPOT WELD", 0.85),
                new NotePattern(@"\bplug\s*weld\b", NoteCategory.Weld, RoutingImpact.AddOperation, RoutingOp.Weld, "F400", "PLUG WELD", 0.85),
                new NotePattern(@"\btack\s*weld\b", NoteCategory.Weld, RoutingImpact.AddOperation, RoutingOp.Weld, "F400", "TACK WELD", 0.85),
                new NotePattern(@"\bfillet\s*weld\b", NoteCategory.Weld, RoutingImpact.AddOperation, RoutingOp.Weld, "F400", "FILLET WELD PER DWG", 0.90),

                // MACHINING / TAPPING
                new NotePattern(@"\btap\s+(\d+[/\-]\d+)", NoteCategory.Machine, RoutingImpact.AddOperation, RoutingOp.Tap, "F220", "TAP {1}", 0.90),
                new NotePattern(@"\bdrill\s+.{1,40}?thru\b", NoteCategory.Machine, RoutingImpact.AddOperation, RoutingOp.Drill, null, "DRILL PER DWG", 0.80),
                new NotePattern(@"\bcountersink\b", NoteCategory.Machine, RoutingImpact.AddOperation, RoutingOp.Machine, null, "COUNTERSINK PER DWG", 0.85),
                new NotePattern(@"\bcounterbore\b", NoteCategory.Machine, RoutingImpact.AddOperation, RoutingOp.Machine, null, "COUNTERBORE PER DWG", 0.85),
                new NotePattern(@"\bream\s+to\b", NoteCategory.Machine, RoutingImpact.AddOperation, RoutingOp.Machine, null, "REAM PER DWG", 0.85),

                // PROCESS CONSTRAINTS
                new NotePattern(@"\bwaterjet\s*only\b", NoteCategory.ProcessConstraint, RoutingImpact.ModifyOperation, RoutingOp.ProcessOverride, "F110", "WATERJET ONLY", 0.95),
                new NotePattern(@"\blaser\s*cut\b", NoteCategory.ProcessConstraint, RoutingImpact.ModifyOperation, RoutingOp.ProcessOverride, "F115", "LASER CUT", 0.85),
                new NotePattern(@"\bplasma\s*cut\b", NoteCategory.ProcessConstraint, RoutingImpact.ModifyOperation, RoutingOp.ProcessOverride, "F120", "PLASMA CUT", 0.85),
                new NotePattern(@"\bdo\s*not\s*(?:laser|burn)\b", NoteCategory.ProcessConstraint, RoutingImpact.ModifyOperation, RoutingOp.ProcessOverride, "F110", "DO NOT LASER - USE WATERJET", 0.95),
                new NotePattern(@"\bflame\s*cut\b", NoteCategory.ProcessConstraint, RoutingImpact.ModifyOperation, RoutingOp.ProcessOverride, "F120", "FLAME CUT", 0.80),

                // INSPECTION
                new NotePattern(@"\binspect\s*(?:per|to|100%|all)\b", NoteCategory.Inspect, RoutingImpact.AddOperation, RoutingOp.Inspect, null, "INSPECT PER DWG", 0.85),
                new NotePattern(@"\bcmm\s*inspect\b", NoteCategory.Inspect, RoutingImpact.AddOperation, RoutingOp.Inspect, null, "CMM INSPECT", 0.90),
                new NotePattern(@"\bfirst\s*article\b", NoteCategory.Inspect, RoutingImpact.AddOperation, RoutingOp.Inspect, null, "FIRST ARTICLE REQUIRED", 0.90),
                new NotePattern(@"\bppap\s*required\b", NoteCategory.Inspect, RoutingImpact.AddOperation, RoutingOp.Inspect, null, "PPAP REQUIRED", 0.90),

                // HARDWARE
                new NotePattern(@"\binstall\s+pem\b", NoteCategory.Hardware, RoutingImpact.AddOperation, RoutingOp.Hardware, null, "INSTALL PEM HARDWARE", 0.90),
                new NotePattern(@"\bpress\s*fit\s*(.{1,60}?)(?:\s*$)", NoteCategory.Hardware, RoutingImpact.AddOperation, RoutingOp.Hardware, null, "PRESS FIT HARDWARE", 0.85),
                new NotePattern(@"\binsert\s+rivet\s*nut\b", NoteCategory.Hardware, RoutingImpact.AddOperation, RoutingOp.Hardware, null, "INSTALL RIVET NUT", 0.85),
                new NotePattern(@"\binstall\s+.{1,40}?insert\b", NoteCategory.Hardware, RoutingImpact.AddOperation, RoutingOp.Hardware, null, "INSTALL INSERT PER DWG", 0.80),
            };
        }
    }

    /// <summary>
    /// A regex pattern mapped to a note category and routing operation.
    /// </summary>
    internal sealed class NotePattern
    {
        public Regex Regex { get; }
        public NoteCategory Category { get; }
        public RoutingImpact Impact { get; }
        public RoutingOp RoutingOp { get; }
        public string WorkCenter { get; }
        public string RoutingNoteTemplate { get; }
        public double Confidence { get; }

        public NotePattern(string pattern, NoteCategory category, RoutingImpact impact,
            RoutingOp routingOp, string workCenter, string routingNoteTemplate, double confidence)
        {
            Regex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.Multiline);
            Category = category;
            Impact = impact;
            RoutingOp = routingOp;
            WorkCenter = workCenter;
            RoutingNoteTemplate = routingNoteTemplate;
            Confidence = confidence;
        }
    }

    /// <summary>
    /// Intermediate match candidate for deduplication of overlapping regex matches.
    /// </summary>
    internal sealed class NoteCandidate
    {
        public int Start;
        public int End;
        public string Text;
        public NotePattern Pattern;
        public int PageNumber;
    }
}
