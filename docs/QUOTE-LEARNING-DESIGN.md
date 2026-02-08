# Past Quote Learning - Design Document

## Problem

Every time an estimator processes a part, they make routing decisions: which work centers,
how much setup/run time, what outside processes. This knowledge lives in their heads and
in completed ERP records, but the system never learns from it. A new estimator has to
build this intuition from scratch.

## Goal

Build a feedback loop where every completed quote teaches the system:
1. **Capture**: When a part finishes the pipeline, save its properties + routing as a record
2. **Index**: Build a searchable database of past quotes keyed by material, thickness, process type, feature count
3. **Suggest**: When processing a new part, find similar past parts and pre-fill routing times
4. **Improve**: Track suggestion accuracy over time and adjust confidence

This is NOT machine learning — it's a structured lookup with similarity scoring.
Simple, explainable, and useful from day one with just a handful of records.

## Architecture

```
Part completes pipeline
    ↓
QuoteRecorder.Capture(partData, properties)
    ↓
QuoteRecord → quoteHistory.json (append)

New part enters pipeline
    ↓
QuoteLookup.FindSimilar(newPartData)
    ↓
Ranked matches by similarity
    ↓
HistoricalSuggestion[] → merge into PropertySuggestionService
```

### New Files

| File | Location | Purpose |
|------|----------|---------|
| QuoteRecord.cs | src/NM.Core/Learning/Models/ | Single completed quote record |
| QuoteHistory.cs | src/NM.Core/Learning/ | JSON-backed quote database with append/query |
| SimilarityScorer.cs | src/NM.Core/Learning/ | Scores similarity between two parts |
| QuoteRecorder.cs | src/NM.Core/Learning/ | Captures completed quotes into history |
| QuoteLookup.cs | src/NM.Core/Learning/ | Finds similar past quotes for suggestions |
| HistoricalSuggestion.cs | src/NM.Core/Learning/Models/ | Suggestion from past quote data |
| QuoteHistoryTests.cs | src/NM.Core.Tests/Learning/ | Tests |
| SimilarityScorerTests.cs | src/NM.Core.Tests/Learning/ | Tests |

### Storage

- **File**: `%APPDATA%\NorthernMfg\quoteHistory.json` (one JSON object per line)
- **Format**: JSON Lines (append-friendly, no need to parse entire file to add)
- **Size estimate**: ~500 bytes per record × 10,000 quotes = ~5MB
- **No external database required**

## QuoteRecord Schema

```json
{
  "id": "guid",
  "timestamp": "2025-01-15T14:30:00Z",
  "fileName": "12345-01.SLDPRT",
  "partNumber": "12345-01",
  "description": "MOUNTING BRACKET",

  "material": "A36",
  "materialType": 0,
  "thickness_in": 0.25,
  "partType": "SheetMetal",

  "features": {
    "bendCount": 4,
    "holeCount": 8,
    "hasTaps": true,
    "hasCountersinks": false,
    "hasHardware": false,
    "flatArea_sqin": 45.2,
    "perimeter_in": 32.1,
    "weight_lb": 2.3
  },

  "routing": {
    "OP20": "F115",
    "OP20_RN": "LASER CUT",
    "F210": "1",
    "F210_RN": "DEBURR",
    "F220": "1",
    "F220_RN": "TAP 1/4-20 (8X)",
    "PressBrake": "Checked",
    "F140_RN": "4 BENDS",
    "OS_WC": "POWDER COAT",
    "OS_RN": "POWDER COAT RAL 9005"
  },

  "times": {
    "OP20_setup_min": 15,
    "OP20_run_min": 2.5,
    "F210_setup_min": 5,
    "F210_run_min": 1.0,
    "F220_setup_min": 10,
    "F220_run_min": 3.0,
    "F140_setup_min": 15,
    "F140_run_min": 4.0
  },

  "source": "Manual",
  "estimator": "TS",
  "confidence": 1.0
}
```

## Similarity Scoring

Match new parts against history using weighted feature comparison:

| Feature | Weight | Scoring |
|---------|--------|---------|
| Material | 25% | Exact match = 1.0, same type = 0.7, different = 0 |
| Thickness | 20% | 1.0 - abs(diff)/max(a,b), min 0 |
| Part Type | 15% | Exact match = 1.0, else 0 |
| Bend Count | 10% | 1.0 - abs(diff)/max(a,b,1) |
| Hole Count | 10% | 1.0 - abs(diff)/max(a,b,1) |
| Flat Area | 10% | 1.0 - abs(diff)/max(a,b,1) |
| Perimeter | 5% | 1.0 - abs(diff)/max(a,b,1) |
| Has Taps | 5% | Match = 1.0, mismatch = 0 |

**Minimum similarity threshold**: 0.60 (below this, don't suggest)

**Confidence scaling**: suggestion confidence = similarity × source confidence

## Suggestion Generation

When similar parts are found:

1. **Setup/Run Times**: Average times from top-3 similar parts, weighted by similarity
2. **Routing Notes**: Use most common routing note from similar parts
3. **Work Centers**: Use most common work center assignment
4. **Outside Process**: If >50% of similar parts had outside process, suggest it

### Example

New part: A36, 0.25" thick, 6 bends, 12 holes, has taps

Top 3 matches:
| Past Part | Similarity | F140 Setup | F140 Run |
|-----------|-----------|------------|----------|
| 12345-01  | 0.92      | 15 min     | 4.0 min  |
| 23456-02  | 0.85      | 12 min     | 3.5 min  |
| 34567-03  | 0.78      | 18 min     | 5.0 min  |

Suggested F140 time: weighted average → ~14.8 min setup, ~4.1 min run

## Integration Points

1. **QuoteRecorder**: Called from `DrawingAnalysisRunner` after properties are written
   - Also hookable from `MainRunner` when standard pipeline completes
2. **QuoteLookup**: Called from `PropertySuggestionService.GeneratePartSuggestions()`
   - Historical suggestions added with source = "Historical ({partNumber})"
   - Lower confidence (scaled by similarity) so drawing-based suggestions take priority
3. **PropertyReviewWizard**: Show "Based on similar part: {partNumber}" in reason column

## Privacy / Data Considerations

- All data stays local (file on disk, not cloud)
- No customer names stored — just part geometry and routing
- User can clear history at any time
- Records auto-expire after configurable period (default: never)

## Implementation Order

1. QuoteRecord + QuoteHistory (storage layer)
2. SimilarityScorer (matching logic)
3. QuoteRecorder (capture from pipeline)
4. QuoteLookup (query similar parts)
5. Integration into PropertySuggestionService
6. UI indicator in PropertyReviewWizard
7. Tests

## Future Enhancements

- **Import from ERP**: Bulk-load historical quotes from Export.prn files
- **Team sharing**: Sync quote history across workstations via shared network drive
- **Analytics dashboard**: "Most quoted materials", "Average setup times by op"
- **Anomaly detection**: Flag quotes that deviate significantly from historical norms
