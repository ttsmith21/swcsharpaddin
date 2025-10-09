# Git Commit Message Template

## Suggested Commit for Handoff

```
feat: Stub tube processing and prepare for handoff

BREAKING: Tube processing is not functional - geometry extraction needs implementation

What's Working:
- Sheet metal conversion pipeline (fully tested with 304L + bend tables)
- Validation pipeline (single-body, multi-body rejection, empty geometry)
- Material selection UI (12 materials)
- Bend table resolution (network paths, UNC mapping, fallbacks)
- Infrastructure (logging, timing, error handling)

What's Stubbed:
- SimpleTubeProcessor.ExtractTubeGeometry() - BLOCKED on cylinder face detection
- All tube services (PipeSchedule, MaterialCode, Cutting, etc.) are complete and ready
- Custom Properties service (Epic 0) - deferred

Handoff Docs:
- HANDOFF_SUMMARY.md - Full status and next steps
- QUICK_STATUS.md - Test scenarios and what works now
- docs/TUBE_HANDOFF.md - Detailed tube debugging guide
- MasterPlan.md - Updated progress tracking

Next Steps for Programmer:
1. Implement SimpleTubeProcessor.ExtractTubeGeometry() (see docs/TUBE_HANDOFF.md)
2. Implement Custom Properties service (Epic 0)
3. Add manufacturing intelligence (weight, cost, bend analysis)

Architecture:
- Clean NM.SwAddin (COM) vs NM.Core (logic) separation maintained
- No DI containers - kept simple per copilot-instructions.md
- All SolidWorks API calls on main thread (STA)

Status: ~75% Phase 1 complete (sheet metal production-ready)
```

## Alternative Shorter Commit

```
feat: Complete sheet metal pipeline, stub tube processing for handoff

Sheet metal conversion is production-ready and tested.
Tube processing stubbed - needs geometry extraction debugging.

See HANDOFF_SUMMARY.md and docs/TUBE_HANDOFF.md for details.
```

## Branch Strategy Recommendation

Current branch: `SinglePartEpic`

Suggested workflow:
1. Commit current work to `SinglePartEpic`
2. Merge to `main` when sheet metal is production-validated
3. Create `feature/tube-geometry` branch for tube work
4. Create `feature/custom-properties` branch for Epic 0

## Files to Include in Commit

```
? Modified:
  - MasterPlan.md (progress update)
  
? New:
  - src/NM.Core/Processing/SimpleTubeProcessor.cs (stubbed)
  - HANDOFF_SUMMARY.md
  - QUICK_STATUS.md
  - docs/TUBE_HANDOFF.md
  - COMMIT_MESSAGE.md (this file)

?? Verify unchanged:
  - src/NM.SwAddin/* (all working code)
  - ErrorHandler.cs, PerformanceTracker.cs (infrastructure)
  - SimpleSheetMetalProcessor.cs (production code)
```

## Pre-Commit Checklist

- [ ] Build succeeds (warnings about locked DLL are normal)
- [ ] SimpleTubeProcessor.cs has no compile errors
- [ ] MasterPlan.md reflects current status
- [ ] All handoff docs are present
- [ ] Sheet metal pipeline still works (smoke test)
- [ ] No sensitive data (passwords, API keys) in commits

## Post-Commit Actions

1. Tag release: `git tag v0.75-sheet-metal-ready`
2. Push to remote: `git push origin SinglePartEpic --tags`
3. Create GitHub issue: "Implement tube geometry extraction"
4. Link handoff docs in issue description

---

*Commit template created: January 2025*
*For handoff to real programmer*
