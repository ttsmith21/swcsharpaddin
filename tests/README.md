# Gold Standard Testing

This directory contains the regression testing framework for validating the SolidWorks Add-in.

## Quick Start

```powershell
# 1. Check what test files you have
.\scripts\check-corpus-coverage.ps1

# 2. Run tests on available files
.\scripts\run-gold-standard-tests.ps1
# Then click "Run QA" in SolidWorks

# 3. Compare results against baseline
.\scripts\compare-qa-results.ps1

# 4. Generate baseline from results (if needed)
.\scripts\create-baseline.ps1 -ResultsFile .\tests\Run_YYYYMMDD\results.json
```

## Current Coverage (Updated 2026-01-29)

| Class | Files | Status |
|-------|-------|--------|
| A - Invalid | 6/6 | ✅ Complete |
| B - Sheet Metal | 3/6 | 🔶 B4-B6 missing |
| C - Tube | 7/7 | ✅ Complete |
| D - Non-Convertible | 0/5 | ❌ Not started |
| E - Edge Cases | 0/6 | ❌ Not started |
| F - Assemblies | 0/8 | ❌ Not started |
| G - Configurations | 0/4 | ❌ Not started |
| H - File Formats | 0/5 | ❌ Not started |

## Directory Structure

```
tests/
├── GoldStandard_Inputs/      # Your test files go here
├── GoldStandard_Baseline/
│   └── manifest.json         # Expected outcomes
├── GoldStandard_Corpus_Spec.md  # Full test spec (47 files)
└── Run_YYYYMMDD_HHMMSS/      # Test run outputs
    ├── Inputs/               # Copied test files
    ├── results.json          # Add-in output
    └── comparison_report.txt # Pass/fail summary
```

## Workflows

### A. Testing with Partial File Sets

You don't need all 47 files to start testing. The harness handles partial sets:

```powershell
# See what's covered
.\scripts\check-corpus-coverage.ps1

# Output:
#   Class A - Invalid/Problem [High]
#     [██████░░░░] 6/6 (100%)
#   Class B - Sheet Metal [High]
#     [████░░░░░░] 4/6 (67%)
#     Missing: B5_BracketWithHoles_14ga_CS.sldprt, B6_TinyFlange_10ga_CS.sldprt
#   ...
```

Tests run on whatever files exist. Missing files are skipped.

### B. Capturing VBA Baseline (Recommended)

To validate C# matches VBA behavior:

```
Step 1: Prepare test files
├── Copy test files to: tests\VBA_Processed\
└── Run VBA macro on all files in SolidWorks

Step 2: Capture VBA output
├── Run: .\scripts\capture-vba-baseline.ps1
├── Click "Run QA" in SolidWorks (reads properties)
└── Results saved to: tests\VBA_Baseline_YYYYMMDD\

Step 3: Generate baseline manifest
├── Run: .\scripts\create-baseline.ps1 -ResultsFile <path>
└── Copy to: tests\GoldStandard_Baseline\manifest.json

Step 4: Test C# against VBA baseline
├── Copy original files to: tests\GoldStandard_Inputs\
├── Run: .\scripts\run-gold-standard-tests.ps1
└── C# results compared against VBA baseline
```

### C. Creating New Baseline (No VBA Reference)

If starting fresh without VBA comparison:

```powershell
# 1. Add test files to GoldStandard_Inputs
# 2. Run tests (will show raw results)
.\scripts\run-gold-standard-tests.ps1

# 3. Review results.json, verify values are correct
# 4. Generate baseline from "known good" run
.\scripts\create-baseline.ps1 -ResultsFile .\tests\Run_YYYYMMDD\results.json

# 5. Edit manifest.json to add tolerances and notes
```

## File Naming Convention

Use the IDs from `GoldStandard_Corpus_Spec.md`:

| Class | Pattern | Example |
|-------|---------|---------|
| A - Invalid | `A#_Name.sldprt` | `A3_MultiBody.sldprt` |
| B - Sheet Metal | `B#_Name_gauge_material.sldprt` | `B1_NativeBracket_14ga_CS.sldprt` |
| C - Tube | `C#_Name_size.sldprt` | `C1_RoundTube_2OD_SCH40.sldprt` |
| D - Other | `D#_Name.sldprt` | `D1_MachinedBlock.sldprt` |
| E - Edge | `E#_Name.sldprt` | `E1_ThickPlate_1in.sldprt` |
| F - Assembly | `F#_Name.sldasm` | `F1_SimpleAssy.sldasm` |
| G - Config | `G#_Name.sldprt` | `G1_MultiConfig.sldprt` |
| H - Import | `H#_Name.ext` | `H1_StepImport.step` |

## Manifest Format

```json
{
  "files": {
    "B1_NativeBracket_14ga_CS.sldprt": {
      "shouldPass": true,
      "expectedClassification": "SheetMetal",
      "expectedThickness_in": 0.0747,
      "expectedBendCount": 1,
      "tolerances": {
        "thickness": 0.001
      }
    },
    "A3_MultiBody.sldprt": {
      "shouldPass": false,
      "expectedFailureReason": "Multi-body"
    }
  }
}
```

## Priority Order

Start with these (19 files, covers core functionality):

1. **Class A** (6 files) - Validates rejection logic
2. **Class B** (6 files) - Sheet metal happy path
3. **Class C** (7 files) - Tube classification

Then expand to:
4. **Class D-E** (11 files) - Edge cases
5. **Class F-G** (12 files) - Assemblies and configs
6. **Class H** (5 files) - Import formats
