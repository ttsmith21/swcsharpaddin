# Goal
Scaffold a clean SOLIDWORKS add?in foundation:
- Project: NM.SwAddin (.NET 4.8, x64)
- References: sldworks, swconst, swpublished (Embed Interop Types = False)
- Implement ISwAddin with ConnectToSW/DisconnectFromSW + SetAddinCallbackInfo2
- One CommandGroup "NM Tools" + 3 commands (Convert, Problems, Fasteners)
- Add Task Pane with CreateTaskpaneView3 (DPI) and a placeholder .NET control
- No icons yet (add later)

# Output
- `SwAddin.cs` (<= ~200 lines)
- Minimal command stubs (call placeholder methods in NM.Core)
- Post-build RegAsm (Framework64) and `[ComRegisterFunction]/[ComUnregisterFunction]` writing HKLM AddIns keys + HKCU AddInsStartup=1
