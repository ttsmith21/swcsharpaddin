# Autonomous Work Plan - Phase 1 Validation

## Current State
- Multi-body validation: ✅ Done
- Material validation: ✅ Done
- Manifest updated: ✅ Done
- Thickness validation: 🔶 Pending (A6_ZeroThickness)
- Comparison tooling: 🔶 Pending

## Proposed Autonomous Work

### 1. Build Comparison Script (Task #3) - LOW RISK
Create a PowerShell script that:
- Reads QA results from `tests/Run_Latest/results.json`
- Compares against `tests/GoldStandard_Baseline/manifest.json`
- Reports mismatches (classification, dimensions, pass/fail status)
- Outputs comparison report

**Risk:** None - this is read-only tooling

### 2. Fix A6_ZeroThickness Validation (Task #5) - MEDIUM RISK
Add validation to fail parts with invalid/zero thickness geometry.

Options:
- A) Add thickness check in MainRunner after classification
- B) Make Generic classification require explicit opt-in (not a fallback)

**Risk:** May affect classification of legitimate Generic parts

### 3. Add Missing Test Part Placeholders - LOW RISK
Create placeholder entries in manifest for missing test files:
- B4-B6 (Sheet Metal edge cases)
- D1-D5 (Non-convertible)
- E1-E6 (Edge cases)

**Risk:** None - documentation only

### 4. Sync QARunner and PartPreflight Paths - LOW RISK
Ensure QARunner uses PartValidationAdapter so all validation
checks are applied consistently (currently bypasses it).

**Risk:** Low - improves consistency

---

## Work I Will NOT Do Without Approval
- Change classification logic (Sheet Metal vs Tube vs Generic routing)
- Modify cost calculations
- Change custom property writes
- Any changes that affect production behavior beyond validation

## Estimated Commits
1. `feat: add QA comparison script`
2. `fix: add thickness validation for invalid geometry`
3. `docs: add missing test file placeholders to manifest`
4. `refactor: route QARunner through PartValidationAdapter`

---

## To Approve
Reply with which items (1-4) I should proceed with, or "all" for everything listed.
