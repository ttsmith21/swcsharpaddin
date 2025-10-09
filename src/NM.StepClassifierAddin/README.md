# NM.StepClassifierAddin

Classifies single-solid STEP bodies into Stick / SheetMetal / Other using a fast?robust decision funnel.

## Build
- Target: .NET Framework 4.8, x64
- Add COM interop references:
  - SolidWorks.Interop.sldworks
  - SolidWorks.Interop.swconst
  - SolidWorks.Interop.swpublished
  - SolidWorks.Interop.gtswutilities (optional; for Utilities thickness analysis)

## Register (Debug)
- Use regasm (Framework64):
  - RegAsm.exe "bin\\x64\\Debug\\NM.StepClassifierAddin.dll" /codebase
- Or set project property RegisterForComInterop=true and run VS as admin.

## Use
- Start SOLIDWORKS, ensure Add-in is loaded (NM Step Classifier).
- Open a part (STEP-imported solid).
- CommandManager group "NM Classifier" ? button "Classify Selected Body".
- A MessageBox shows Stick / SheetMetal / Other.
- Logs written to %TEMP%\\NM.StepClassifier\\log.txt.

## Notes
- Utilities must be installed for ThicknessAnalyzer to run; otherwise Phase 3 is skipped.
- The add-in avoids adding features; all geometry ops use transient bodies and modeler APIs.
