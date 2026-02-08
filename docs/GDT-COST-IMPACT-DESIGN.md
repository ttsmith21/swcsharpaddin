# GD&T Cost Impact Flagging - Design Document

## Problem

Engineering drawings contain GD&T (Geometric Dimensioning & Tolerancing) callouts that
significantly impact manufacturing cost — tight tolerances require slower feeds, special
tooling, CMM inspection, or even different processes entirely. Today, estimators manually
scan drawings for these, which is slow and error-prone. Missing a tight tolerance on a
quote means eating the cost difference.

## Goal

Automatically extract GD&T callouts from PDF drawings and flag those that drive cost:
- True position callouts with tight tolerance
- Surface finish requirements (Ra/RMS values)
- Flatness/parallelism/perpendicularity beyond standard
- Profile tolerances
- Concentricity/runout requirements

For each flagged callout, estimate the cost impact category (Low/Medium/High) and
suggest routing additions (CMM inspect, grinding, fixture, etc.).

## Architecture

```
DrawingText → GdtExtractor → GdtCallout[] → CostImpactAnalyzer → CostFlag[]
                                                    ↓
                                            RoutingHint[] (inspect, grind, etc.)
```

### New Files

| File | Location | Purpose |
|------|----------|---------|
| GdtExtractor.cs | src/NM.Core/Pdf/ | Regex-based GD&T symbol/callout extraction |
| CostImpactAnalyzer.cs | src/NM.Core/Pdf/ | Evaluates extracted GD&T for cost drivers |
| CostFlag.cs | src/NM.Core/Pdf/Models/ | Cost impact result model |
| GdtExtractorTests.cs | src/NM.Core.Tests/Pdf/ | Tests |
| CostImpactAnalyzerTests.cs | src/NM.Core.Tests/Pdf/ | Tests |

### Integration Points

1. **PdfDrawingAnalyzer.AnalyzeCore()** — add GD&T extraction step after note extraction
2. **DrawingData.GdtCallouts** — already exists (empty list, ready to populate)
3. **ReconciliationEngine** — add cost flags as a new field on ReconciliationResult
4. **PropertyReviewWizard** — add a "Cost Flags" section showing high-risk callouts

## GD&T Patterns to Extract

### Tolerance Callouts (Text-Based)
```
±0.005          → Bilateral tolerance
+0.002/-0.000   → Unilateral tolerance
0.001 TRUE POS  → True position
```

### Surface Finish
```
Ra 32           → Surface roughness (microinches)
RMS 63          → Surface roughness (RMS)
125/            → Surface finish symbol value
```

### GD&T Symbols (Unicode or Text Abbreviations)
```
⌀ 0.005 TRUE POSITION   → True position
⏥ 0.002                 → Flatness
∥ 0.003                  → Parallelism
⊥ 0.002                  → Perpendicularity
◎ 0.010                  → Concentricity
↗ 0.005                  → Runout
⌓ 0.003                  → Profile of a line
```

Since PDF text extraction often loses special symbols, also match text equivalents:
```
TRUE POSITION 0.005
FLATNESS 0.002
PARALLELISM 0.003
PERPENDICULARITY 0.002
CONCENTRICITY 0.010
TOTAL RUNOUT 0.005
PROFILE 0.003
```

## Cost Impact Thresholds

| Feature | Standard | Moderate | Tight | Very Tight |
|---------|----------|----------|-------|------------|
| True Position | > 0.010" | 0.005-0.010" | 0.002-0.005" | < 0.002" |
| Flatness | > 0.010" | 0.005-0.010" | 0.002-0.005" | < 0.002" |
| Surface Finish (Ra) | > 125 | 63-125 | 32-63 | < 32 |
| General Tolerance | ±0.010" | ±0.005" | ±0.002" | ±0.001" |
| Angular | ±1° | ±0.5° | ±0.25° | ±5 min |

### Cost Impact Categories

| Impact | What It Means | Routing Addition |
|--------|---------------|-----------------|
| Low | Standard machining can hold this | None |
| Medium | May need slower feeds or better tooling | Note on routing |
| High | Needs CMM, grinding, or special fixture | Add inspect op |
| Critical | May affect process selection entirely | Flag for review |

## Routing Suggestions from GD&T

| Condition | Suggested Action |
|-----------|-----------------|
| Any tolerance < 0.002" | Add CMM INSPECT routing note |
| Surface finish Ra < 32 | Add GRINDING operation |
| True position < 0.005" | Add "FIXTURE REQUIRED" note |
| Profile < 0.003" | Add CMM INSPECT + fixture note |
| Multiple tight tolerances | Flag: "MULTIPLE TIGHT TOLERANCES - REVIEW PRICING" |

## Implementation Order

1. GdtExtractor — regex patterns for all GD&T types
2. CostImpactAnalyzer — threshold-based evaluation
3. Integration into PdfDrawingAnalyzer pipeline
4. Add cost flags to ReconciliationResult
5. Display in PropertyReviewWizard
6. Tests

## Dependencies

- Builds on Phase 1 (PdfTextExtractor) for text access
- Builds on existing GdtCallout model (already defined, currently empty)
- Integrates with existing RoutingHint pipeline
