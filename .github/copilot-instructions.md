# Northern Manufacturing — SOLIDWORKS Add?in (VBA ? C#) — Copilot Rules

## Mission
Port a large, partially?debugged **VBA macro** into a **C# SOLIDWORKS add?in** that is simple, reliable, and easy to extend. Build a clear foundation now; add power later without rewrites.

## Environment (authoritative)
- Visual Studio on Windows
- Target **.NET Framework 4.8**; build **x64** (no .NET Core/5+/6+).
- Interop references from **install_dir\api\redist**:
  - `SolidWorks.Interop.sldworks.dll`
  - `SolidWorks.Interop.swconst.dll`
  - `SolidWorks.Interop.swpublished.dll`  (ISwAddin lives here)
  - `SolidWorksTools.dll` **only** if we add CommandManager/Task Pane icons
- Implement `ISwAddin`. In `ConnectToSW`, call `ISldWorks.SetAddinCallbackInfo2`.

## Foundation (Goldilocks, not minimal)
- **Solution layout**
  - `src/NM.SwAddin/` — COM add?in (UI, events, SW wiring), thin glue only
  - `src/NM.Core/` — pure C# logic (no COM types)
  - `tests/NM.Core.Tests/` — unit tests for `NM.Core` only
- **Keep it simple**
  - No DI containers, mediators, or async. Small classes, early returns.
  - Use interfaces (`ISldWorks`, `IModelDoc2`, …) and enums (no magic numbers).
- **Units**
  - Internal units = meters; convert only for display.

## Migration principles (VBA ? C#)
1) **Port intent, not shape.** Replace late binding with strong types.
2) **Thin glue.** Put SOLIDWORKS API calls in `NM.SwAddin`; keep algorithms in `NM.Core`.
3) **Happy path first.** Add `// TODO(vNext)` markers for edge cases that can wait.

## Initial command surface
- **Convert to Sheet Metal** — primary workflow; clear success/fail messages
- **List Problem Parts** — finds parts with no solid body or failed conversions
- **Analyze Fasteners** — simple heuristic (small body + helical curves)

## Common pitfalls to avoid (hard rules)
- **Targeting** — Always `.NET 4.8`, `x64`. Do not use “Any CPU”.
- **Registration** — Use `[ComRegisterFunction]/[ComUnregisterFunction]` to write:
  - `HKLM\Software\SOLIDWORKS\AddIns\{GUID}` ? Default (1/0), Title, Description
  - `HKCU\Software\SOLIDWORKS\AddInsStartup\{GUID}` ? 1 to autoload
  - Use **Framework64** `RegAsm.exe` in Debug (`/codebase` OK for dev; sign for Release).
- **Task Pane (DPI aware)** — Use `ISldWorks.CreateTaskpaneView3` and host with `ITaskpaneView.DisplayWindowFromHandlex64`. Provide six icon sizes (20/32/40/64/96/128 px).
- **CommandManager** — When button definitions change, assign a **new UserID** or set `IgnorePreviousVersion = true` to avoid duplicate/ghost groups.
- **Selections** — Validate selection counts and types using `ISelectionMgr` (`GetSelectedObjectCount2`, `GetSelectedObjectType3`, `GetSelectedObject6`).
- **Feature edits** — Always follow `GetDefinition` ? modify data ? `IFeature.ModifyDefinition`.
- **Threading** — Treat SOLIDWORKS COM as STA/main?thread only. Do not call SW API on background threads. Background tasks may compute data, but all COM calls must marshal to the main thread.
- **COM lifetime** — Do **not** blanket?call `Marshal.ReleaseComObject`. Use it **only** in measured, high?volume loops for short?lived RCWs you created. Never release top?level singletons like `ISldWorks`.

## No Complexity Drift (scope guardrails)
- Feature tasks: ? ~3 files and ? ~200 added lines unless explicitly requested.
- YAGNI: no DI frameworks, no plugin buses, no reflection helpers, no custom exception hierarchies, no telemetry systems.

## Debugging Protocol (regression?proof)
When the request is to **debug/fix**:
- Persona = **Code Surgeon**.
- Modify **only** the function or block named/pasted by the user.
- Return a **Minimal Viable Fix (MVF)** plus an **Impact Note**:
  1) Root cause (1–2 sentences); 2) What changed (exact file/method); 3) Why it’s safe.
- If a broader change is required, **stop** and list the dependencies; wait for approval.

## Testing & validation
- Unit tests live in `NM.Core.Tests` for pure logic (no COM).
- Every Copilot answer that changes code should include a short **Manual Smoke**:
  - Open a simple imported part; run **Convert to Sheet Metal**.
  - Expect sheet?metal feature or one clear failure message.
  - Open an assembly; run **List Problem Parts** and **Analyze Fasteners**; verify output.

## Deployment (dev workflow)
- Debug ? Start external program: `SLDWORKS.exe`.
- Post?build (Debug): `RegAsm.exe "$(TargetPath)" /codebase` from Framework64.
- Release later: MSI/WiX; set HKLM AddIns and (optionally) HKCU AddInsStartup.

## How to ask Copilot (templates)
- **Scaffold:** “Create `ISwAddin` with `SetAddinCallbackInfo2`, one CommandGroup ‘NM Tools’, and commands (Convert, Problems, Fasteners). Use CreateTaskpaneView3; no icons yet.”
- **Port a module:** “Port attached VBA into `NM.Core/SheetMetalConverter` (no COM types) with 5 tests. Keep happy path; add TODOs for edge cases.”
- **Debug mode:** “DEBUG this method only. Explain the null ref and propose the smallest fix; don’t change signatures or other files.”
