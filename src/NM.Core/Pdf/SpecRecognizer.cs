using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using NM.Core.Pdf.Models;

namespace NM.Core.Pdf
{
    /// <summary>
    /// Recognizes manufacturing specifications, standards, and codes in drawing text.
    /// Covers ASTM, AMS, MIL-SPEC, AWS, ASME, SAE, and common OEM callouts.
    /// Each recognized spec maps to a routing impact (material confirmation, outside process, etc.).
    /// </summary>
    public sealed class SpecRecognizer
    {
        private static readonly SpecEntry[] _specs = BuildDatabase();

        /// <summary>
        /// Master regex that matches any known spec identifier in text.
        /// Compiled once for performance across many pages.
        /// </summary>
        private static readonly Regex _masterPattern = BuildMasterPattern();

        /// <summary>
        /// Scans text for known specification references.
        /// Returns structured matches with routing implications.
        /// </summary>
        public List<SpecMatch> Recognize(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new List<SpecMatch>();

            var results = new List<SpecMatch>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // First pass: master pattern to find all spec-like references quickly
            var candidates = _masterPattern.Matches(text);
            foreach (Match candidate in candidates)
            {
                string raw = candidate.Value.Trim();
                if (seen.Contains(raw)) continue;

                // Second pass: match against specific entries for classification
                foreach (var spec in _specs)
                {
                    if (!spec.Pattern.IsMatch(raw)) continue;
                    seen.Add(raw);

                    results.Add(new SpecMatch
                    {
                        RawText = raw,
                        SpecId = spec.SpecId,
                        FullName = spec.FullName,
                        Category = spec.Category,
                        RoutingOp = spec.RoutingOp,
                        WorkCenter = spec.WorkCenter,
                        RoutingNote = spec.RoutingNote,
                        Confidence = spec.Confidence
                    });
                    break;
                }
            }

            return results;
        }

        /// <summary>
        /// Converts spec matches into routing hints that integrate with the existing pipeline.
        /// </summary>
        public List<RoutingHint> ToRoutingHints(List<SpecMatch> matches)
        {
            if (matches == null || matches.Count == 0)
                return new List<RoutingHint>();

            var hints = new List<RoutingHint>();
            var seenOps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var match in matches)
            {
                if (match.RoutingOp == null)
                    continue;

                // Deduplicate: don't create two routing hints for the same operation
                string opKey = $"{match.RoutingOp}:{match.WorkCenter ?? "null"}";
                if (!seenOps.Add(opKey))
                    continue;

                hints.Add(new RoutingHint
                {
                    Operation = match.RoutingOp.Value,
                    WorkCenter = match.WorkCenter,
                    NoteText = match.RoutingNote ?? $"PER {match.SpecId}",
                    SourceNote = match.RawText,
                    Confidence = match.Confidence
                });
            }

            return hints;
        }

        /// <summary>
        /// Converts spec matches into drawing notes for integration with the note pipeline.
        /// </summary>
        public List<DrawingNote> ToDrawingNotes(List<SpecMatch> matches, int pageNumber = 1)
        {
            if (matches == null || matches.Count == 0)
                return new List<DrawingNote>();

            return matches.Select(m => new DrawingNote
            {
                Text = $"{m.SpecId}: {m.FullName}",
                Category = MapCategory(m.Category),
                Impact = m.RoutingOp.HasValue ? RoutingImpact.AddOperation : RoutingImpact.Informational,
                Confidence = m.Confidence,
                PageNumber = pageNumber
            }).ToList();
        }

        private static NoteCategory MapCategory(SpecCategory category)
        {
            switch (category)
            {
                case SpecCategory.Material: return NoteCategory.Material;
                case SpecCategory.Welding: return NoteCategory.Weld;
                case SpecCategory.Coating: return NoteCategory.Finish;
                case SpecCategory.Plating: return NoteCategory.Finish;
                case SpecCategory.HeatTreat: return NoteCategory.HeatTreat;
                case SpecCategory.Inspection: return NoteCategory.Inspect;
                case SpecCategory.SurfaceFinish: return NoteCategory.Finish;
                case SpecCategory.Process: return NoteCategory.Machine;
                case SpecCategory.Quality: return NoteCategory.Inspect;
                case SpecCategory.Testing: return NoteCategory.Inspect;
                case SpecCategory.Controlled: return NoteCategory.General;
                default: return NoteCategory.General;
            }
        }

        private static Regex BuildMasterPattern()
        {
            // Catches anything that looks like a spec reference
            // ASTM A-36, AMS 4027, MIL-PRF-22750, AWS D1.1, ASME Y14.5, SAE J429, etc.
            return new Regex(
                @"\b(?:" +
                    @"ASTM\s*[A-Z][\-\s]?\d+" +
                    @"|AMS[\-\s]?\d{4}" +
                    @"|AMS[\-\s][A-Z][\-\s]?\d+" +
                    @"|MIL[\-\s][A-Z]{1,4}[\-\s]\d+" +
                    @"|MIL[\-\s]DTL[\-\s]\d+" +
                    @"|MIL[\-\s]STD[\-\s]\d+" +
                    @"|AWS\s*[A-Z]\d+(?:\.\d+)?" +
                    @"|ASME\s*[A-Z]\d+(?:\.\d+)?" +
                    @"|SAE\s*[A-Z]?\d+" +
                    @"|QQ[\-\s][A-Z][\-\s]\d+" +
                    @"|NADCAP" +
                    @"|AS\s*9100" +
                    @"|AS\s*9102" +
                    @"|ISO\s*9001" +
                    @"|ITAR" +
                    @"|DFARS" +
                    @"|CUI\b" +
                    @"|NIST\s*800[\-\s]\d+" +
                    @"|AMS[\-\s]QQ[\-\s][A-Z][\-\s]\d+" +
                @")" +
                @"(?:[\-\s]*(?:CLASS|TYPE|GRADE|COND|GR)\s*[A-Z0-9]{1,4})?",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);
        }

        // =====================================================================
        // SPEC DATABASE
        // Each entry: pattern, specId, fullName, category, routingOp, workCenter, routingNote, confidence
        // =====================================================================
        private static SpecEntry[] BuildDatabase()
        {
            return new[]
            {
                // ===== MATERIAL SPECIFICATIONS (ASTM Structural Steel) =====
                S(@"ASTM\s*A[\-\s]?36",  "ASTM A36",  "Carbon Structural Steel",          SpecCategory.Material, null, null, null, 0.95),
                S(@"ASTM\s*A[\-\s]?500", "ASTM A500", "Structural Tubing (Cold-Formed)",   SpecCategory.Material, null, null, null, 0.95),
                S(@"ASTM\s*A[\-\s]?513", "ASTM A513", "Electric-Resistance-Welded Tubing", SpecCategory.Material, null, null, null, 0.95),
                S(@"ASTM\s*A[\-\s]?514", "ASTM A514", "High-Yield Quenched & Tempered Plate", SpecCategory.Material, null, null, null, 0.90),
                S(@"ASTM\s*A[\-\s]?516", "ASTM A516", "Pressure Vessel Plate",             SpecCategory.Material, null, null, null, 0.90),
                S(@"ASTM\s*A[\-\s]?529", "ASTM A529", "High-Strength Carbon-Manganese Steel", SpecCategory.Material, null, null, null, 0.90),
                S(@"ASTM\s*A[\-\s]?572", "ASTM A572", "High-Strength Low-Alloy Steel",     SpecCategory.Material, null, null, null, 0.95),
                S(@"ASTM\s*A[\-\s]?588", "ASTM A588", "Weathering Steel (Corten)",          SpecCategory.Material, null, null, null, 0.90),
                S(@"ASTM\s*A[\-\s]?53",  "ASTM A53",  "Pipe (Black/Galvanized)",           SpecCategory.Material, null, null, null, 0.90),

                // ===== MATERIAL SPECIFICATIONS (ASTM Stainless) =====
                S(@"ASTM\s*A[\-\s]?240", "ASTM A240", "Stainless Steel Plate/Sheet",       SpecCategory.Material, null, null, null, 0.95),
                S(@"ASTM\s*A[\-\s]?276", "ASTM A276", "Stainless Steel Bar",               SpecCategory.Material, null, null, null, 0.90),
                S(@"ASTM\s*A[\-\s]?312", "ASTM A312", "Stainless Steel Pipe",              SpecCategory.Material, null, null, null, 0.90),
                S(@"ASTM\s*A[\-\s]?269", "ASTM A269", "Stainless Steel Tubing",            SpecCategory.Material, null, null, null, 0.90),
                S(@"ASTM\s*A[\-\s]?554", "ASTM A554", "Stainless Welded Mechanical Tubing", SpecCategory.Material, null, null, null, 0.90),

                // ===== MATERIAL SPECIFICATIONS (ASTM Aluminum) =====
                S(@"ASTM\s*B[\-\s]?209", "ASTM B209", "Aluminum Sheet/Plate",              SpecCategory.Material, null, null, null, 0.90),
                S(@"ASTM\s*B[\-\s]?221", "ASTM B221", "Aluminum Bar/Rod/Wire/Shape",       SpecCategory.Material, null, null, null, 0.90),
                S(@"ASTM\s*B[\-\s]?241", "ASTM B241", "Aluminum Seamless Pipe/Tube",       SpecCategory.Material, null, null, null, 0.90),

                // ===== MATERIAL SPECIFICATIONS (AMS Aerospace) =====
                S(@"AMS[\-\s]?4027",  "AMS 4027",  "6061-T6 Aluminum Sheet",               SpecCategory.Material, null, null, null, 0.90),
                S(@"AMS[\-\s]?4041",  "AMS 4041",  "2024-T3 Aluminum Sheet",               SpecCategory.Material, null, null, null, 0.90),
                S(@"AMS[\-\s]?4044",  "AMS 4044",  "2024-T4 Aluminum Sheet",               SpecCategory.Material, null, null, null, 0.90),
                S(@"AMS[\-\s]?4911",  "AMS 4911",  "Titanium 6Al-4V Sheet/Plate",          SpecCategory.Material, null, null, null, 0.90),
                S(@"AMS[\-\s]?5510",  "AMS 5510",  "304 Stainless Sheet",                  SpecCategory.Material, null, null, null, 0.90),
                S(@"AMS[\-\s]?5524",  "AMS 5524",  "321 Stainless Sheet",                  SpecCategory.Material, null, null, null, 0.90),
                S(@"AMS[\-\s]?5596",  "AMS 5596",  "Inconel 718 Sheet",                    SpecCategory.Material, null, null, null, 0.90),
                S(@"AMS[\-\s]?6350",  "AMS 6350",  "4130 Normalized Steel",                SpecCategory.Material, null, null, null, 0.90),

                // ===== WELDING CODES =====
                S(@"AWS\s*D1[\.\s]?1",  "AWS D1.1",  "Structural Welding - Steel",          SpecCategory.Welding, RoutingOp.Weld, "F400", "WELD PER AWS D1.1", 0.95),
                S(@"AWS\s*D1[\.\s]?2",  "AWS D1.2",  "Structural Welding - Aluminum",       SpecCategory.Welding, RoutingOp.Weld, "F400", "WELD PER AWS D1.2", 0.95),
                S(@"AWS\s*D1[\.\s]?3",  "AWS D1.3",  "Structural Welding - Sheet Steel",    SpecCategory.Welding, RoutingOp.Weld, "F400", "WELD PER AWS D1.3", 0.95),
                S(@"AWS\s*D1[\.\s]?6",  "AWS D1.6",  "Structural Welding - Stainless",      SpecCategory.Welding, RoutingOp.Weld, "F400", "WELD PER AWS D1.6", 0.95),
                S(@"AWS\s*D17[\.\s]?1", "AWS D17.1", "Aerospace Fusion Welding",            SpecCategory.Welding, RoutingOp.Weld, "F400", "WELD PER AWS D17.1 (AEROSPACE)", 0.95),
                S(@"AWS\s*D9[\.\s]?1",  "AWS D9.1",  "Sheet Metal Welding",                 SpecCategory.Welding, RoutingOp.Weld, "F400", "WELD PER AWS D9.1", 0.90),

                // ===== COATING / PAINT SPECS =====
                S(@"MIL[\-\s]PRF[\-\s]22750",  "MIL-PRF-22750",  "Epoxy Primer (High-Solids)", SpecCategory.Coating, RoutingOp.OutsideProcess, null, "PRIME PER MIL-PRF-22750", 0.90),
                S(@"MIL[\-\s]PRF[\-\s]85285",  "MIL-PRF-85285",  "Polyurethane Topcoat",       SpecCategory.Coating, RoutingOp.OutsideProcess, null, "PAINT PER MIL-PRF-85285", 0.90),
                S(@"MIL[\-\s]DTL[\-\s]53039",  "MIL-DTL-53039",  "CARC Epoxy Primer",          SpecCategory.Coating, RoutingOp.OutsideProcess, null, "PRIME PER MIL-DTL-53039 (CARC)", 0.90),
                S(@"MIL[\-\s]PRF[\-\s]23377",  "MIL-PRF-23377",  "Epoxy Primer",               SpecCategory.Coating, RoutingOp.OutsideProcess, null, "PRIME PER MIL-PRF-23377", 0.90),
                S(@"AMS[\-\s]C[\-\s]27725",    "AMS-C-27725",    "Powder Coating",             SpecCategory.Coating, RoutingOp.OutsideProcess, null, "POWDER COAT PER AMS-C-27725", 0.90),

                // ===== PLATING SPECS =====
                S(@"ASTM\s*B[\-\s]?633",       "ASTM B633",       "Zinc Electroplating",        SpecCategory.Plating, RoutingOp.OutsideProcess, null, "ZINC PLATE PER ASTM B633", 0.90),
                S(@"ASTM\s*B[\-\s]?456",       "ASTM B456",       "Nickel/Chrome Plating",      SpecCategory.Plating, RoutingOp.OutsideProcess, null, "NICKEL/CHROME PLATE PER ASTM B456", 0.90),
                S(@"ASTM\s*B[\-\s]?488",       "ASTM B488",       "Gold Electroplating",        SpecCategory.Plating, RoutingOp.OutsideProcess, null, "GOLD PLATE PER ASTM B488", 0.85),
                S(@"ASTM\s*B[\-\s]?733",       "ASTM B733",       "Electroless Nickel",         SpecCategory.Plating, RoutingOp.OutsideProcess, null, "ELECTROLESS NICKEL PER ASTM B733", 0.90),
                S(@"AMS[\-\s]QQ[\-\s]N[\-\s]290", "AMS-QQ-N-290", "Nickel Plating",             SpecCategory.Plating, RoutingOp.OutsideProcess, null, "NICKEL PLATE PER AMS-QQ-N-290", 0.90),
                S(@"QQ[\-\s]N[\-\s]290",        "QQ-N-290",       "Nickel Plating (Legacy)",     SpecCategory.Plating, RoutingOp.OutsideProcess, null, "NICKEL PLATE PER QQ-N-290", 0.85),
                S(@"MIL[\-\s]DTL[\-\s]5541",   "MIL-DTL-5541",   "Chemical Film (Chem Film / Alodine)", SpecCategory.Plating, RoutingOp.OutsideProcess, null, "CHEM FILM PER MIL-DTL-5541", 0.90),
                S(@"MIL[\-\s]DTL[\-\s]13924",  "MIL-DTL-13924",  "Black Oxide",                SpecCategory.Plating, RoutingOp.OutsideProcess, null, "BLACK OXIDE PER MIL-DTL-13924", 0.90),
                S(@"AMS[\-\s]2700",            "AMS 2700",        "Passivation (Stainless)",    SpecCategory.Plating, RoutingOp.OutsideProcess, null, "PASSIVATE PER AMS 2700", 0.90),
                S(@"ASTM\s*A[\-\s]?967",       "ASTM A967",       "Passivation (Chemical)",     SpecCategory.Plating, RoutingOp.OutsideProcess, null, "PASSIVATE PER ASTM A967", 0.90),

                // ===== HEAT TREAT SPECS =====
                S(@"AMS[\-\s]2759",  "AMS 2759",  "Heat Treatment of Steel Parts",     SpecCategory.HeatTreat, RoutingOp.OutsideProcess, null, "HEAT TREAT PER AMS 2759", 0.90),
                S(@"AMS[\-\s]H[\-\s]6875",  "AMS-H-6875",  "Heat Treatment of Stainless",  SpecCategory.HeatTreat, RoutingOp.OutsideProcess, null, "HEAT TREAT PER AMS-H-6875", 0.90),
                S(@"AMS[\-\s]2750",  "AMS 2750",  "Pyrometry (Furnace Calibration)",   SpecCategory.HeatTreat, null, null, null, 0.85),

                // ===== SURFACE FINISH / TREATMENT SPECS =====
                S(@"AMS[\-\s]2430",  "AMS 2430",  "Shot Peening",                      SpecCategory.SurfaceFinish, RoutingOp.OutsideProcess, null, "SHOT PEEN PER AMS 2430", 0.90),
                S(@"AMS[\-\s]2431",  "AMS 2431",  "Shot Peening (Computer Monitored)", SpecCategory.SurfaceFinish, RoutingOp.OutsideProcess, null, "SHOT PEEN PER AMS 2431", 0.90),
                S(@"MIL[\-\s]STD[\-\s]171",  "MIL-STD-171",  "Surface Finishing",      SpecCategory.SurfaceFinish, RoutingOp.OutsideProcess, null, "FINISH PER MIL-STD-171", 0.85),
                S(@"AMS[\-\s]2470",  "AMS 2470",  "Anodize Type I (Chromic)",           SpecCategory.SurfaceFinish, RoutingOp.OutsideProcess, null, "ANODIZE TYPE I PER AMS 2470", 0.90),
                S(@"AMS[\-\s]2471",  "AMS 2471",  "Anodize Type II (Sulfuric)",         SpecCategory.SurfaceFinish, RoutingOp.OutsideProcess, null, "ANODIZE TYPE II PER AMS 2471", 0.90),
                S(@"AMS[\-\s]2472",  "AMS 2472",  "Anodize Type III (Hard)",            SpecCategory.SurfaceFinish, RoutingOp.OutsideProcess, null, "HARD ANODIZE PER AMS 2472", 0.90),
                S(@"MIL[\-\s]A[\-\s]8625",  "MIL-A-8625",  "Anodic Coatings for Aluminum", SpecCategory.SurfaceFinish, RoutingOp.OutsideProcess, null, "ANODIZE PER MIL-A-8625", 0.90),

                // ===== INSPECTION / TESTING SPECS =====
                S(@"ASME\s*Y14[\.\s]?5",  "ASME Y14.5",  "Dimensioning & Tolerancing (GD&T)", SpecCategory.Inspection, null, null, null, 0.90),
                S(@"AS\s*9102",            "AS 9102",     "First Article Inspection",           SpecCategory.Inspection, RoutingOp.Inspect, null, "FIRST ARTICLE PER AS 9102", 0.95),
                S(@"ASTM\s*E[\-\s]?1444",  "ASTM E1444",  "Magnetic Particle Inspection",      SpecCategory.Testing, RoutingOp.Inspect, null, "MAG PARTICLE INSPECT PER ASTM E1444", 0.90),
                S(@"ASTM\s*E[\-\s]?1417",  "ASTM E1417",  "Liquid Penetrant Inspection",       SpecCategory.Testing, RoutingOp.Inspect, null, "LPI PER ASTM E1417", 0.90),
                S(@"ASTM\s*E[\-\s]?94",    "ASTM E94",    "Radiographic Examination",           SpecCategory.Testing, RoutingOp.Inspect, null, "RADIOGRAPHIC INSPECT PER ASTM E94", 0.90),
                S(@"ASTM\s*E[\-\s]?164",   "ASTM E164",   "Ultrasonic Contact Examination",     SpecCategory.Testing, RoutingOp.Inspect, null, "UT INSPECT PER ASTM E164", 0.90),

                // ===== QUALITY MANAGEMENT =====
                S(@"AS\s*9100",    "AS 9100",    "Aerospace Quality Management System",    SpecCategory.Quality, null, null, null, 0.90),
                S(@"ISO\s*9001",   "ISO 9001",   "Quality Management System",              SpecCategory.Quality, null, null, null, 0.85),
                S(@"NADCAP",       "NADCAP",     "National Aerospace & Defense Contractors Accreditation Program", SpecCategory.Quality, null, null, null, 0.90),

                // ===== FASTENER / HARDWARE SPECS =====
                S(@"SAE\s*J429",    "SAE J429",    "Mechanical Properties of Bolts",       SpecCategory.Material, null, null, null, 0.80),
                S(@"ASTM\s*A[\-\s]?193",  "ASTM A193",  "High-Temp Bolting Material",      SpecCategory.Material, null, null, null, 0.80),
                S(@"ASTM\s*A[\-\s]?194",  "ASTM A194",  "High-Temp Nut Material",          SpecCategory.Material, null, null, null, 0.80),
                S(@"ASTM\s*F[\-\s]?3125", "ASTM F3125", "High-Strength Structural Bolts",  SpecCategory.Material, null, null, null, 0.80),

                // ===== CONTROLLED INFORMATION FLAGS =====
                S(@"\bITAR\b",         "ITAR",         "International Traffic in Arms Regulations",       SpecCategory.Controlled, null, null, null, 0.95),
                S(@"\bDFARS\b",        "DFARS",        "Defense Federal Acquisition Regulation Supplement", SpecCategory.Controlled, null, null, null, 0.90),
                S(@"\bCUI\b",          "CUI",          "Controlled Unclassified Information",              SpecCategory.Controlled, null, null, null, 0.90),
                S(@"NIST\s*800[\-\s]171", "NIST 800-171", "Protecting CUI in Nonfederal Systems",          SpecCategory.Controlled, null, null, null, 0.90),

                // ===== PROCESS SPECS =====
                S(@"AMS[\-\s]2175",  "AMS 2175",  "Classification of Castings",           SpecCategory.Process, null, null, null, 0.85),
                S(@"AMS[\-\s]2644",  "AMS 2644",  "Fluorescent Penetrant Inspection",     SpecCategory.Testing, RoutingOp.Inspect, null, "FPI PER AMS 2644", 0.90),
                S(@"AMS[\-\s]2645",  "AMS 2645",  "Fluorescent Penetrant Inspection (Type 1)", SpecCategory.Testing, RoutingOp.Inspect, null, "FPI TYPE 1 PER AMS 2645", 0.90),
            };
        }

        /// <summary>
        /// Helper to create a SpecEntry with less boilerplate.
        /// </summary>
        private static SpecEntry S(string pattern, string specId, string fullName,
            SpecCategory category, RoutingOp? routingOp, string workCenter,
            string routingNote, double confidence)
        {
            return new SpecEntry
            {
                Pattern = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled),
                SpecId = specId,
                FullName = fullName,
                Category = category,
                RoutingOp = routingOp,
                WorkCenter = workCenter,
                RoutingNote = routingNote,
                Confidence = confidence
            };
        }

        /// <summary>
        /// Returns the count of specs in the database.
        /// </summary>
        public static int DatabaseSize => _specs.Length;
    }

    /// <summary>
    /// A single entry in the spec recognition database.
    /// </summary>
    internal sealed class SpecEntry
    {
        public Regex Pattern { get; set; }
        public string SpecId { get; set; }
        public string FullName { get; set; }
        public SpecCategory Category { get; set; }
        public RoutingOp? RoutingOp { get; set; }
        public string WorkCenter { get; set; }
        public string RoutingNote { get; set; }
        public double Confidence { get; set; }
    }

    /// <summary>
    /// Result of recognizing a specification in drawing text.
    /// </summary>
    public sealed class SpecMatch
    {
        /// <summary>The raw text that matched (e.g., "ASTM A-36").</summary>
        public string RawText { get; set; }

        /// <summary>Normalized spec identifier (e.g., "ASTM A36").</summary>
        public string SpecId { get; set; }

        /// <summary>Human-readable full name (e.g., "Carbon Structural Steel").</summary>
        public string FullName { get; set; }

        /// <summary>Category of the specification.</summary>
        public SpecCategory Category { get; set; }

        /// <summary>Routing operation implied by this spec (null if informational only).</summary>
        public RoutingOp? RoutingOp { get; set; }

        /// <summary>Work center code if applicable.</summary>
        public string WorkCenter { get; set; }

        /// <summary>Routing note text for the ERP.</summary>
        public string RoutingNote { get; set; }

        /// <summary>Confidence in the match (0.0-1.0).</summary>
        public double Confidence { get; set; }
    }

    /// <summary>
    /// Categories of manufacturing specifications.
    /// </summary>
    public enum SpecCategory
    {
        Material,
        Welding,
        Coating,
        Plating,
        HeatTreat,
        SurfaceFinish,
        Inspection,
        Testing,
        Quality,
        Process,
        Controlled
    }
}
