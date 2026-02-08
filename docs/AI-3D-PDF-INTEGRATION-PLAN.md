# AI-Powered 3D Model + PDF Drawing Integration

## Vision: The Manufacturing Intelligence Engine

**Problem:** Job shops receive RFQs as a mix of 3D models (.sldprt, .step) and 2D PDF drawings. Today, estimators manually cross-reference both, type part numbers into the ERP, guess routings from experience, and hope nothing falls through the cracks. Competitors like Paperless Parts require uploading files to a cloud platform, losing native SolidWorks metadata. None of them generate actual shop-floor routings — they only estimate cost.

**Solution:** An AI-powered system that lives *inside* SolidWorks, analyzes both the 3D model and its companion PDF drawing simultaneously, cross-validates them, extracts every piece of manufacturing-relevant data, and generates complete ERP-ready routings with notes — all without leaving the SolidWorks environment.

**Tagline:** *"The estimator's brain, running at machine speed."*

---

## Competitive Differentiation

| Capability | Paperless Parts | aPriori | SecturaSOFT | **NM Intelligence Engine** |
|-----------|----------------|---------|-------------|---------------------------|
| Native SolidWorks integration | No (STEP upload) | No (STEP upload) | No (DXF import) | **Yes — reads .sldprt, .slddrw, custom props** |
| PDF drawing analysis | Wingman AI (cloud) | No | No | **Built-in, offline-capable** |
| 3D + 2D cross-validation | No | No | No | **Yes — reconciles model vs drawing** |
| Generates actual routings | No (cost only) | No (cost only) | No (cost only) | **Yes — OP20-OP60 with work centers** |
| ERP export with routing notes | Limited | No | Limited | **Yes — Import.prn with RT + RN sections** |
| Sheet metal + tube in one flow | No | No | No | **Yes** |
| Works offline | No (cloud) | No (cloud) | Desktop | **Yes — desktop, optional AI cloud** |
| Feedback loop (actual vs quoted) | No | No | No | **Planned** |
| Cost | ~$500+/mo SaaS | Enterprise ($$$) | Mid-range | **One-time / included** |

### The Key Insight Competitors Miss

Cloud platforms **strip away** the richest data source (native CAD properties) by requiring neutral format uploads. We **start** with the richest possible data from the 3D model, then **augment** it with AI-extracted drawing intelligence. This dual-source approach catches errors that neither source alone reveals.

---

## Architecture

### System Overview

```
┌─────────────────────────────────────────────────────────────────────┐
│                    NM Manufacturing Intelligence Engine              │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  ┌──────────────┐    ┌──────────────┐    ┌───────────────────────┐  │
│  │  3D Model     │    │  PDF Drawing  │    │  AI Analysis Engine   │  │
│  │  Analyzer     │    │  Analyzer     │    │  (Claude Vision API)  │  │
│  │              │    │              │    │                       │  │
│  │ • Geometry    │    │ • Text OCR    │    │ • Drawing interpret.  │  │
│  │ • Properties  │    │ • Title block │    │ • Note extraction     │  │
│  │ • Sheet metal │    │ • Notes       │    │ • GD&T parsing        │  │
│  │ • Tube data   │    │ • BOM table   │    │ • Routing suggestions │  │
│  │ • Material    │    │ • Rev block   │    │ • Discrepancy detect  │  │
│  └──────┬───────┘    └──────┬───────┘    └───────────┬───────────┘  │
│         │                   │                        │              │
│         └───────────────────┴────────────────────────┘              │
│                             │                                       │
│                    ┌────────▼────────┐                              │
│                    │  Reconciliation  │                              │
│                    │  Engine          │                              │
│                    │                  │                              │
│                    │ • Cross-validate │                              │
│                    │ • Fill gaps      │                              │
│                    │ • Flag conflicts │                              │
│                    │ • Confidence %   │                              │
│                    └────────┬────────┘                              │
│                             │                                       │
│         ┌───────────────────┼───────────────────┐                  │
│         ▼                   ▼                   ▼                  │
│  ┌──────────────┐  ┌───────────────┐  ┌─────────────────┐        │
│  │ Property      │  │ Routing       │  │ ERP Export      │        │
│  │ Writeback     │  │ Generator     │  │ (Import.prn)    │        │
│  │              │  │              │  │                 │        │
│  │ • Part number │  │ • OP20-OP60   │  │ • Item Master   │        │
│  │ • Description │  │ • Work centers│  │ • Routing       │        │
│  │ • Revision    │  │ • Setup/Run   │  │ • Routing Notes │        │
│  │ • Material    │  │ • Notes       │  │ • BOM           │        │
│  │ • File rename │  │ • Special ops │  │ • Material      │        │
│  └──────────────┘  └───────────────┘  └─────────────────┘        │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

### Data Flow

```
1. User opens part in SolidWorks (or batch processes folder)
2. System locates companion PDF drawing (same name, or user-selected)
3. PARALLEL ANALYSIS:
   a. 3D Model → existing pipeline (geometry, sheet metal, tube, properties)
   b. PDF Drawing → text extraction (PdfPig) + AI vision analysis (Claude API)
4. Reconciliation engine merges both data sources:
   - If 3D says material = "1018 CRS" and PDF says "ASTM A36" → FLAG CONFLICT
   - If 3D has no description but PDF title block says "MOUNTING BRACKET" → FILL GAP
   - If PDF notes say "BREAK ALL EDGES" but no deburr in routing → ADD OP30
5. User reviews AI suggestions in a reconciliation dialog
6. Approved changes → property writeback + routing generation + ERP export
```

---

## Module Design

### Module 1: PDF Drawing Analyzer (`NM.Core.Pdf`)

**Pure C# — no COM dependencies. Fully unit-testable.**

```
src/NM.Core/Pdf/
├── PdfDrawingAnalyzer.cs       # Orchestrates text + AI analysis
├── PdfTextExtractor.cs         # PdfPig-based text extraction
├── TitleBlockParser.cs         # Regex/pattern matching for title blocks
├── DrawingNoteExtractor.cs     # Extracts general notes, callouts
├── BomTableExtractor.cs        # Extracts BOM tables from PDF
└── Models/
    ├── DrawingData.cs           # Central DTO (mirrors PartData for drawings)
    ├── TitleBlockInfo.cs        # Part number, material, rev, date, drawn by
    ├── DrawingNote.cs           # Note text, category, confidence
    ├── BomEntry.cs              # Item no, part number, description, qty
    └── DimensionInfo.cs         # Extracted dimensions with confidence
```

**Key class: `DrawingData`** (the PDF equivalent of `PartData`)

```csharp
public sealed class DrawingData
{
    // Identity (from title block)
    public string PartNumber { get; set; }
    public string Description { get; set; }
    public string Revision { get; set; }
    public string Material { get; set; }
    public string DrawnBy { get; set; }
    public DateTime? DrawingDate { get; set; }

    // Geometry hints (from dimensions/notes)
    public double? Thickness_in { get; set; }
    public double? OverallLength_in { get; set; }
    public double? OverallWidth_in { get; set; }

    // Manufacturing notes
    public List<DrawingNote> Notes { get; set; } = new List<DrawingNote>();
    public List<DrawingNote> GdtCallouts { get; set; } = new List<DrawingNote>();

    // BOM (if assembly drawing)
    public List<BomEntry> BomEntries { get; set; } = new List<BomEntry>();

    // Routing hints extracted from notes
    public List<RoutingHint> RoutingHints { get; set; } = new List<RoutingHint>();

    // Source tracking
    public string SourcePdfPath { get; set; }
    public int PageCount { get; set; }
    public AnalysisMethod Method { get; set; } // TextOnly, VisionAI, Hybrid
    public double OverallConfidence { get; set; }
}
```

### Module 2: AI Vision Service (`NM.Core.AI`)

**Abstracted AI provider — works with Claude, GPT-4o, or offline fallback.**

```
src/NM.Core/AI/
├── IDrawingVisionService.cs     # Interface for AI analysis
├── ClaudeVisionService.cs       # Claude API implementation
├── OfflineVisionService.cs      # Regex/heuristic fallback (no API needed)
├── VisionPrompts.cs             # Structured prompts for drawing analysis
├── VisionResponseParser.cs      # Parses AI JSON responses
└── Models/
    ├── VisionAnalysisResult.cs  # Structured AI output
    ├── VisionConfig.cs          # API keys, model selection, DPI settings
    └── FieldConfidence.cs       # Per-field confidence tracking
```

**Interface design:**

```csharp
public interface IDrawingVisionService
{
    /// <summary>
    /// Analyzes a PDF drawing page using AI vision.
    /// Returns null if AI is unavailable (offline mode).
    /// </summary>
    Task<VisionAnalysisResult> AnalyzePageAsync(byte[] pngImageBytes, VisionContext context);

    /// <summary>
    /// Checks if the AI service is available and configured.
    /// </summary>
    bool IsAvailable { get; }

    /// <summary>
    /// Estimated cost per page in USD.
    /// </summary>
    decimal EstimatedCostPerPage { get; }
}
```

**Tiered analysis strategy:**

```
Tier 1 (Free, Offline): PdfPig text extraction + regex parsing
  → Works for digitally-created PDFs with machine-readable text
  → Extracts: part number, material, revision, basic notes
  → Confidence: ~70-85% for clean title blocks

Tier 2 (AI-Assisted): Claude Vision API on title block region only
  → For scanned PDFs or when Tier 1 confidence < 60%
  → Crops title block area, sends only that region (~small image)
  → Cost: ~$0.001 per page (tiny image)
  → Confidence: ~90-95% for title block data

Tier 3 (Full AI): Claude Vision API on full drawing
  → For complex drawings with notes, GD&T, routing callouts
  → Sends full page at 200 DPI
  → Cost: ~$0.004 per page
  → Confidence: 85-95% for notes, 70-80% for GD&T
```

### Module 3: Reconciliation Engine (`NM.Core.Reconciliation`)

**The brain — merges 3D + 2D data and resolves conflicts.**

```
src/NM.Core/Reconciliation/
├── ReconciliationEngine.cs      # Main merge logic
├── ConflictResolver.cs          # Rules for resolving conflicts
├── GapFiller.cs                 # Fills missing data from alternate source
├── RoutingHintInterpreter.cs    # Converts drawing notes → routing ops
├── FileRenameService.cs         # Generates rename suggestions
└── Models/
    ├── ReconciliationResult.cs  # Merged data + conflicts + suggestions
    ├── DataConflict.cs          # Describes a 3D vs 2D disagreement
    ├── GapFill.cs               # Describes data filled from alternate source
    └── RoutingSuggestion.cs     # Suggested routing modification
```

**Key class: `ReconciliationEngine`**

```csharp
public sealed class ReconciliationEngine
{
    /// <summary>
    /// Merges 3D model data with PDF drawing data.
    /// Returns reconciled result with conflicts, gap fills, and suggestions.
    /// </summary>
    public ReconciliationResult Reconcile(PartData modelData, DrawingData drawingData)
    {
        var result = new ReconciliationResult();

        // 1. CROSS-VALIDATE: Compare overlapping fields
        ComparePartNumbers(modelData, drawingData, result);
        CompareMaterials(modelData, drawingData, result);
        CompareThickness(modelData, drawingData, result);
        CompareDescriptions(modelData, drawingData, result);

        // 2. GAP-FILL: Use drawing data to fill 3D model gaps
        FillMissingDescription(modelData, drawingData, result);
        FillMissingRevision(modelData, drawingData, result);
        FillMissingMaterial(modelData, drawingData, result);

        // 3. ROUTING INTELLIGENCE: Convert notes → operations
        InterpretRoutingNotes(drawingData.Notes, result);

        // 4. FILE RENAME: Suggest renames based on drawing part number
        SuggestFileRename(modelData, drawingData, result);

        return result;
    }
}
```

**Conflict detection examples:**

| 3D Model Says | PDF Drawing Says | Action |
|---------------|------------------|--------|
| No description | "MOUNTING BRACKET" | **Fill** description from drawing |
| Material: blank | "MATERIAL: 304 SS" | **Fill** material, map to OptiMaterial |
| Thickness: 0.125" | Thickness: 0.125" | **Confirm** — increase confidence |
| Thickness: 0.125" | Thickness: 0.250" | **Conflict** — flag for human review |
| Part number: blank | "PN: 12345-A REV C" | **Fill** part number + revision |
| Filename: Part1.sldprt | "PN: BRACKET-ASM-001" | **Suggest** file rename |
| No deburr routing | Note: "BREAK ALL EDGES" | **Add** OP30 deburr to routing |
| No paint routing | Note: "PAINT RED PER SPEC" | **Add** outside processing note |
| Laser routing | Note: "WATERJET ONLY" | **Override** OP20 → F110 waterjet |

### Module 4: Routing Note Intelligence (`NM.Core.Manufacturing.RoutingNotes`)

**Pattern-matches drawing notes to manufacturing operations.**

```
src/NM.Core/Manufacturing/RoutingNotes/
├── NoteClassifier.cs            # Classifies notes into routing categories
├── OperationMatcher.cs          # Maps notes → specific work centers
├── NotePatterns.cs              # Regex patterns for common manufacturing notes
└── Models/
    └── RoutingHint.cs           # Suggested operation from a note
```

**Note → Operation mapping (extensible rule set):**

```csharp
// Built-in patterns — shops can add custom patterns via config
private static readonly NotePattern[] Patterns = new[]
{
    // DEBURR / EDGE BREAK
    new NotePattern(@"break\s*(all)?\s*edges", RoutingOp.Deburr, "F210"),
    new NotePattern(@"deburr\s*(all)?", RoutingOp.Deburr, "F210"),
    new NotePattern(@"remove\s*(all)?\s*burrs", RoutingOp.Deburr, "F210"),
    new NotePattern(@"tumble\s*deburr", RoutingOp.Deburr, "F210"),

    // SURFACE FINISH
    new NotePattern(@"paint\s+(.+)", RoutingOp.OutsideProcess, null, "PAINT: {1}"),
    new NotePattern(@"powder\s*coat", RoutingOp.OutsideProcess, null, "POWDER COAT"),
    new NotePattern(@"anodize", RoutingOp.OutsideProcess, null, "ANODIZE"),
    new NotePattern(@"galvanize", RoutingOp.OutsideProcess, null, "GALVANIZE"),
    new NotePattern(@"zinc\s*plate", RoutingOp.OutsideProcess, null, "ZINC PLATE"),
    new NotePattern(@"chrome\s*plate", RoutingOp.OutsideProcess, null, "CHROME PLATE"),
    new NotePattern(@"black\s*oxide", RoutingOp.OutsideProcess, null, "BLACK OXIDE"),

    // HEAT TREAT
    new NotePattern(@"heat\s*treat", RoutingOp.OutsideProcess, null, "HEAT TREAT"),
    new NotePattern(@"stress\s*reliev", RoutingOp.OutsideProcess, null, "STRESS RELIEVE"),
    new NotePattern(@"harden\s*(to|per)", RoutingOp.OutsideProcess, null, "HARDEN"),
    new NotePattern(@"rc\s*\d{2}", RoutingOp.OutsideProcess, null, "HARDEN TO {0}"),

    // WELDING
    new NotePattern(@"weld\s*(all|per)", RoutingOp.Weld, "F400", "WELD PER DWG"),
    new NotePattern(@"mig\s*weld", RoutingOp.Weld, "F400", "MIG WELD"),
    new NotePattern(@"tig\s*weld", RoutingOp.Weld, "F400", "TIG WELD"),
    new NotePattern(@"spot\s*weld", RoutingOp.Weld, "F400", "SPOT WELD"),

    // MACHINING
    new NotePattern(@"tap\s+(\d+[/-]\d+)", RoutingOp.Tap, "F220", "TAP {1}"),
    new NotePattern(@"drill\s+.*thru", RoutingOp.Drill, null, "DRILL PER DWG"),
    new NotePattern(@"countersink", RoutingOp.Machine, null, "COUNTERSINK"),
    new NotePattern(@"counterbore", RoutingOp.Machine, null, "COUNTERBORE"),
    new NotePattern(@"ream\s+to", RoutingOp.Machine, null, "REAM PER DWG"),

    // PROCESS CONSTRAINTS
    new NotePattern(@"waterjet\s*only", RoutingOp.ProcessOverride, "F110"),
    new NotePattern(@"laser\s*cut", RoutingOp.ProcessOverride, "F115"),
    new NotePattern(@"plasma\s*cut", RoutingOp.ProcessOverride, "F120"),
    new NotePattern(@"do\s*not\s*(laser|burn)", RoutingOp.ProcessOverride, "F110"),

    // INSPECTION
    new NotePattern(@"inspect\s*(per|to|100%)", RoutingOp.Inspect, null, "INSPECT PER DWG"),
    new NotePattern(@"cmm\s*inspect", RoutingOp.Inspect, null, "CMM INSPECT"),
    new NotePattern(@"first\s*article", RoutingOp.Inspect, null, "FIRST ARTICLE REQUIRED"),

    // HARDWARE
    new NotePattern(@"install\s+pem", RoutingOp.Hardware, null, "INSTALL PEM HARDWARE"),
    new NotePattern(@"press\s*fit\s*(.+)", RoutingOp.Hardware, null, "PRESS FIT {1}"),
};
```

### Module 5: File Rename Service (`NM.Core.FileManagement`)

**Generates safe file rename operations based on reconciled data.**

```
src/NM.Core/FileManagement/
├── FileRenameService.cs         # Generates rename plans
├── NamingConvention.cs          # Configurable naming rules
└── Models/
    └── RenamePlan.cs            # Old path → new path + validation
```

```csharp
public class RenamePlan
{
    public string OldPath { get; set; }        // C:\Parts\Part1.sldprt
    public string NewPath { get; set; }        // C:\Parts\12345-A_BRACKET.sldprt
    public string OldDrawingPath { get; set; } // C:\Parts\Part1.slddrw
    public string NewDrawingPath { get; set; } // C:\Parts\12345-A_BRACKET.slddrw
    public string Reason { get; set; }         // "Part number extracted from PDF title block"
    public double Confidence { get; set; }     // 0.95
    public bool RequiresUserApproval { get; set; } // true (always for renames)

    // References that need updating (assemblies pointing to old filename)
    public List<string> AffectedAssemblies { get; set; }
}
```

### Module 6: UI — Reconciliation Wizard (`NM.SwAddin.UI`)

**WinForms dialog that presents AI findings for human review.**

```
src/NM.SwAddin/UI/
├── ReconciliationWizard.cs      # Main wizard form
├── ConflictPanel.cs             # Shows 3D vs 2D conflicts with accept/reject
├── GapFillPanel.cs              # Shows suggested gap fills with approve/skip
├── RoutingSuggestionPanel.cs    # Shows routing changes with accept/modify
├── FileRenamePanel.cs           # Shows rename suggestions with approve/skip
├── PdfPreviewPanel.cs           # Shows PDF page with highlighted regions
└── AiConfigDialog.cs            # API key setup, tier selection, cost tracking
```

**Wizard flow:**

```
Step 1: SELECT PDF
  ├── Auto-detect companion PDF (same name as .sldprt)
  ├── Or browse to select manually
  └── Show PDF preview alongside 3D model

Step 2: AI ANALYSIS (progress bar)
  ├── Text extraction (instant)
  ├── Title block parsing (instant)
  ├── Vision AI analysis (2-5 seconds, if enabled)
  └── Show: "Analyzed 3 pages, found 12 data points, cost: $0.01"

Step 3: REVIEW CONFLICTS (if any)
  ├── Side-by-side: "3D says X, Drawing says Y"
  ├── User picks: Accept 3D | Accept Drawing | Manual Override
  └── Confidence indicators per field

Step 4: REVIEW GAP FILLS
  ├── "Description not set in model. Drawing says: MOUNTING BRACKET"
  ├── "Revision not set. Drawing says: REV C"
  ├── User: Accept | Skip | Edit
  └── Batch accept all with confidence > 90%

Step 5: REVIEW ROUTING SUGGESTIONS
  ├── "Note: 'BREAK ALL EDGES' → Add OP30 Deburr (F210)"
  ├── "Note: 'PAINT RED PER MIL-PRF-22750' → Add Outside Process"
  ├── "Note: 'WATERJET ONLY' → Change OP20 from F115 to F110"
  └── User: Accept | Skip | Modify

Step 6: FILE RENAME (if applicable)
  ├── "Rename Part1.sldprt → 12345-A_BRACKET.sldprt?"
  ├── Show affected assemblies
  └── User: Accept | Skip

Step 7: APPLY
  ├── Write custom properties to SolidWorks model
  ├── Update routing operations
  ├── Rename files (if approved)
  ├── Generate ERP export
  └── Show summary: "Updated 8 properties, added 2 routing ops, renamed 1 file"
```

---

## Implementation Phases

### Phase 1: Foundation — PDF Text Extraction + Title Block Parsing
**Scope:** Extract machine-readable text from PDFs, parse title blocks offline.

| Task | Location | Dependencies |
|------|----------|-------------|
| Add PdfPig NuGet package | `packages.config` or NuGet | None |
| Add PDFtoImage NuGet package | `packages.config` or NuGet | None |
| Create `DrawingData` model | `src/NM.Core/Pdf/Models/` | None |
| Create `PdfTextExtractor` | `src/NM.Core/Pdf/` | PdfPig |
| Create `TitleBlockParser` | `src/NM.Core/Pdf/` | PdfTextExtractor |
| Create `DrawingNoteExtractor` | `src/NM.Core/Pdf/` | PdfTextExtractor |
| Unit tests for title block parsing | `src/NM.Core.Tests/Pdf/` | Sample PDFs |
| Wire into MainRunner as optional step | `src/NM.SwAddin/Pipeline/` | DrawingData model |

**Deliverable:** Given a digital PDF, extract part number, material, revision, description, and general notes with ~80% accuracy. Zero API cost.

### Phase 2: AI Vision — Claude API Integration
**Scope:** Send drawing images to Claude for intelligent extraction.

| Task | Location | Dependencies |
|------|----------|-------------|
| Create `IDrawingVisionService` interface | `src/NM.Core/AI/` | None |
| Implement `ClaudeVisionService` | `src/NM.Core/AI/` | Anthropic.SDK or HttpClient |
| Implement `OfflineVisionService` (fallback) | `src/NM.Core/AI/` | None |
| Create structured prompts for drawing analysis | `src/NM.Core/AI/` | None |
| Create `VisionResponseParser` | `src/NM.Core/AI/` | None |
| Add API key configuration UI | `src/NM.SwAddin/UI/` | None |
| Cost tracking per session | `src/NM.Core/AI/` | None |
| Integration tests with sample drawings | `src/NM.Core.Tests/AI/` | API key |

**Deliverable:** Scanned and complex PDFs analyzed by AI with 90%+ accuracy on title block data. ~$0.004 per page.

### Phase 3: Reconciliation — 3D + 2D Merge
**Scope:** Cross-validate and merge model data with drawing data.

| Task | Location | Dependencies |
|------|----------|-------------|
| Create `ReconciliationEngine` | `src/NM.Core/Reconciliation/` | PartData, DrawingData |
| Implement conflict detection rules | `src/NM.Core/Reconciliation/` | None |
| Implement gap-fill logic | `src/NM.Core/Reconciliation/` | None |
| Material name normalization | `src/NM.Core/Materials/` | OptiMaterial service |
| Unit tests for reconciliation | `src/NM.Core.Tests/` | None |

**Deliverable:** System identifies conflicts, fills gaps, and produces a reconciled dataset with per-field confidence scores.

### Phase 4: Routing Intelligence — Notes → Operations
**Scope:** Interpret manufacturing notes and generate routing suggestions.

| Task | Location | Dependencies |
|------|----------|-------------|
| Create `NoteClassifier` with pattern library | `src/NM.Core/Manufacturing/RoutingNotes/` | None |
| Create `OperationMatcher` | `src/NM.Core/Manufacturing/RoutingNotes/` | Work center definitions |
| Map common notes to operations | `src/NM.Core/Manufacturing/RoutingNotes/` | None |
| Add outside process detection | `src/NM.Core/Manufacturing/RoutingNotes/` | None |
| Wire into ERP export (RN section) | `src/NM.Core/Export/` | ErpExportDataBuilder |
| Unit tests with real-world note examples | `src/NM.Core.Tests/` | None |

**Deliverable:** Drawing notes like "BREAK ALL EDGES" automatically generate OP30 deburr routing step. Notes like "PAINT RED" generate outside process routing notes.

### Phase 5: Property Writeback + File Rename
**Scope:** Write reconciled data back to SolidWorks and rename files.

| Task | Location | Dependencies |
|------|----------|-------------|
| Extend property writeback for drawing-sourced data | `src/NM.SwAddin/Properties/` | CustomPropertiesService |
| Create `FileRenameService` | `src/NM.Core/FileManagement/` | None |
| Handle assembly reference updates on rename | `src/NM.SwAddin/` | SolidWorks API |
| Add revision tracking | `src/NM.SwAddin/Properties/` | None |
| Integration tests | `src/NM.Core.Tests/` | SolidWorks |

**Deliverable:** AI-extracted descriptions, part numbers, and revisions written to SolidWorks custom properties. Files renamed to match drawing part numbers.

### Phase 6: UI — Reconciliation Wizard
**Scope:** Human-in-the-loop review interface.

| Task | Location | Dependencies |
|------|----------|-------------|
| Create `ReconciliationWizard` form | `src/NM.SwAddin/UI/` | Phases 1-5 |
| PDF preview panel | `src/NM.SwAddin/UI/` | PDFtoImage |
| Conflict review panel | `src/NM.SwAddin/UI/` | ReconciliationResult |
| Batch accept/reject controls | `src/NM.SwAddin/UI/` | None |
| Progress and cost display | `src/NM.SwAddin/UI/` | AI cost tracking |
| Wire into add-in command menu | `src/NM.SwAddin/` | SwAddin.cs |

**Deliverable:** Complete wizard flow from PDF selection → AI analysis → review → apply.

### Phase 7: Batch Processing + Feedback Loop
**Scope:** Process entire folders and learn from corrections.

| Task | Location | Dependencies |
|------|----------|-------------|
| Extend FolderProcessor for PDF pairing | `src/NM.SwAddin/Processing/` | Phase 1 |
| Batch reconciliation mode | `src/NM.Core/Reconciliation/` | Phase 3 |
| Correction logging (what user changed) | `src/NM.Core/` | Phase 6 |
| Accuracy metrics dashboard | `src/NM.SwAddin/UI/` | Correction log |
| Custom pattern import (shop-specific notes) | `src/NM.Core/Manufacturing/RoutingNotes/` | Phase 4 |

**Deliverable:** Process 100 parts with PDFs in batch. Track accuracy. Import shop-specific note patterns.

---

## Technology Stack

### NuGet Packages Required

| Package | Version | License | Purpose | .NET 4.8 Compatible |
|---------|---------|---------|---------|---------------------|
| `UglyToad.PdfPig` | 0.1.13+ | Apache 2.0 | PDF text extraction | Yes (.NET 4.5+) |
| `PDFtoImage` | latest | MIT | PDF → PNG conversion | Yes (.NET 4.7.1+) |
| `Newtonsoft.Json` | 13.x | MIT | JSON parsing (already in project) | Yes |
| `Anthropic.SDK` | 2.x | MIT | Claude API client | Yes (netstandard2.0) |

**Total new dependencies: 3** (PdfPig, PDFtoImage, Anthropic.SDK). All MIT/Apache licensed. All compatible with .NET Framework 4.8.

### API Cost Estimates

| Scenario | Pages/Month | Tier | Monthly Cost |
|----------|-------------|------|-------------|
| Small shop, manual | 50 | Tier 2 (title block only) | ~$0.05 |
| Medium shop, mixed | 500 | Tier 2 + some Tier 3 | ~$1-2 |
| Large shop, batch | 5,000 | Tier 3 (full analysis) | ~$20 |
| Offline only | Any | Tier 1 (text extraction) | $0 |

Compare to Paperless Parts: $500+/month minimum.

---

## Prompt Engineering — Drawing Analysis

### Title Block Extraction Prompt (Tier 2)

```
Analyze this engineering drawing title block and extract information as JSON.
Return ONLY valid JSON, no other text.

{
  "part_number": "",
  "description": "",
  "revision": "",
  "material": "",
  "finish": "",
  "drawn_by": "",
  "date": "",
  "scale": "",
  "sheet": "",
  "tolerance_general": "",
  "third_angle_projection": true/false
}

Rules:
- Use empty string "" for fields you cannot determine
- For material, include the full specification (e.g., "ASTM A36", "304 STAINLESS STEEL")
- For revision, include the letter/number only (e.g., "C", "3")
- For part_number, exclude revision suffixes
```

### Full Drawing Analysis Prompt (Tier 3)

```
You are a manufacturing engineer analyzing an engineering drawing.
Extract ALL manufacturing-relevant information as JSON.

{
  "title_block": {
    "part_number": "",
    "description": "",
    "revision": "",
    "material": "",
    "finish": ""
  },
  "dimensions": {
    "overall_length_inches": null,
    "overall_width_inches": null,
    "thickness_inches": null,
    "units": "inches"
  },
  "manufacturing_notes": [
    {
      "text": "exact note text",
      "category": "deburr|finish|heat_treat|weld|machine|inspect|hardware|process|general",
      "routing_impact": "add_operation|modify_operation|informational|constraint"
    }
  ],
  "gdt_callouts": [
    {
      "type": "flatness|parallelism|perpendicularity|position|profile|runout|concentricity",
      "tolerance": "",
      "datum_references": [],
      "feature_description": ""
    }
  ],
  "hole_table": {
    "count": 0,
    "tapped_holes": [],
    "through_holes": []
  },
  "bend_notes": {
    "bend_radius": "",
    "k_factor": "",
    "bend_direction": ""
  },
  "special_requirements": []
}

Rules:
- Only include data you can clearly read from the drawing
- For dimensions, convert to inches if in metric
- For notes, preserve exact wording
- Categorize each note by its manufacturing impact
```

---

## Competitive Moat — What Makes This Unbeatable

### 1. Native CAD Intelligence
Cloud platforms work with dead geometry (STEP files). We work with **living SolidWorks documents** — custom properties, configurations, bend tables, material databases, feature trees. This is 10x richer input than any competitor receives.

### 2. Dual-Source Validation
No competitor cross-validates 3D models against 2D drawings. This catches errors that slip through single-source analysis:
- Drawing updated but model not rebuilt
- Model material changed but drawing not reprinted
- Title block says "REV C" but model is still REV B

### 3. Routing Generation, Not Just Costing
Paperless Parts tells you "this will cost $47.32." We generate the actual **routing** — OP20 laser on F115, setup 0.15hr, run 0.08hr, OP30 deburr, OP40 brake 3 bends — that flows directly to the shop floor via ERP import. This is the gap between "quoting software" and "manufacturing automation."

### 4. Offline-First, AI-Optional
The system works **without** internet or API keys. Text extraction and pattern matching handle 70%+ of drawings. AI is an accelerator, not a requirement. No monthly SaaS bill. No vendor lock-in.

### 5. Shop-Specific Learning
The note pattern library is extensible. Northern Manufacturing's specific callouts ("NM SPEC 101", "ROUTE TO PAINT LINE 2") become first-class routing rules. Over time, the system learns the shop's language.

### 6. Pennies vs Dollars
At ~$0.004 per page for full AI analysis, processing 1,000 drawings costs $4. Paperless Parts charges $500+/month. The cost difference is three orders of magnitude.

---

## Risk Mitigation

| Risk | Mitigation |
|------|-----------|
| AI hallucination on critical data (part numbers) | Human review required for all changes. Confidence scores gate auto-acceptance. |
| Scanned PDFs with poor quality | Tier 1 (text-only) fails gracefully, Tier 3 handles via vision. User warned of low confidence. |
| API rate limits / outages | Offline Tier 1 always works. Failed AI calls don't block the pipeline. |
| File rename breaks assemblies | Rename service identifies affected assemblies. Requires explicit user approval. |
| Drawing-model version mismatch | Conflict detection flags discrepancies. Timestamp comparison warns if drawing is older. |
| Non-English drawings | Claude handles multiple languages. Note patterns extensible per locale. |

---

## Success Metrics

| Metric | Target | Measurement |
|--------|--------|-------------|
| Title block extraction accuracy | >90% on first pass | Compare against manual entry |
| Time saved per part (with PDF) | >5 minutes | Before/after timing study |
| Routing note detection rate | >80% of actionable notes | Manual review of missed notes |
| False positive rate (wrong data applied) | <2% | Track corrections in wizard |
| User acceptance of suggestions | >70% batch accept rate | Log accept/reject in wizard |
| Cost per part analyzed | <$0.01 | Track API spend |

---

## Summary

This isn't a quoting tool — it's a **manufacturing intelligence engine** that sits inside the engineer's native environment (SolidWorks), understands both the 3D geometry and the 2D intent, and produces shop-floor-ready output. The combination of native CAD access + AI-powered drawing analysis + automated routing generation is something no competitor offers today.

The phased approach means value is delivered from Phase 1 (free offline text extraction) while the AI capabilities build progressively. By Phase 4, the system is generating routing operations from drawing notes — a capability that doesn't exist anywhere in the market.
