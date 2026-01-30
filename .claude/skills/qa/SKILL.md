---
name: qa
description: Run headless QA tests to validate code changes. Use after making changes to verify correctness.
---

# QA Validation Skill

This skill runs the headless QA test suite to validate code changes against the gold standard test inputs.

## What This Does

1. Builds the main project (swcsharpaddin.dll)
2. Builds NM.BatchRunner.exe
3. Launches SolidWorks headlessly
4. Runs all 16 gold standard test parts
5. Reports pass/fail results
6. Closes SolidWorks automatically

## Commands to Execute

Run these commands in sequence:

```powershell
# Step 0: Check if SolidWorks is running (DLL will be locked)
$swRunning = Get-Process -Name SLDWORKS -ErrorAction SilentlyContinue
if ($swRunning) {
    Write-Error "SolidWorks is running! Close it first or the DLL will be locked."
    # Do NOT proceed - build will fail or use stale DLL
}

# Step 1: Build main project
powershell -ExecutionPolicy Bypass -File "C:\Users\tsmith\source\repos\swcsharpaddin\scripts\build-and-test.ps1" -SkipClean

# Step 2: Verify DLL was updated (not locked/stale)
$dll = Get-Item "C:\Users\tsmith\source\repos\swcsharpaddin\bin\Debug\swcsharpaddin.dll"
$age = (Get-Date) - $dll.LastWriteTime
if ($age.TotalMinutes -gt 2) {
    Write-Warning "DLL is $($age.TotalMinutes.ToString('F0')) minutes old - may be stale/locked!"
}

# Step 3: Build BatchRunner (if build succeeded)
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe" "C:\Users\tsmith\source\repos\swcsharpaddin\src\NM.BatchRunner\NM.BatchRunner.csproj" /p:Configuration=Debug /v:minimal

# Step 4: Run QA tests
& "C:\Users\tsmith\source\repos\swcsharpaddin\src\NM.BatchRunner\bin\Debug\NM.BatchRunner.exe" --qa
```

## Pre-flight Checks

Before running QA, verify:
1. **SolidWorks is NOT running** - DLL is locked while SW is open
2. **Build succeeds** - Check for "BUILD: SUCCESS" output
3. **DLL timestamp is fresh** - Should be within last 2 minutes

## Interpreting Results

**Exit codes:**
- `0` = All tests passed
- `1` = Some tests failed or errors occurred

**Output files:**
- `C:\Temp\nm_qa_summary.txt` - Quick summary
- `C:\Temp\nm_qa_config.json` - Test configuration
- `tests/Run_Latest/results.json` - Detailed results

**Test categories:**
- `A*` series: Validation edge cases (should fail gracefully)
- `B*` series: Sheet metal parts (should process successfully)
- `C*` series: Structural shapes/tubes (should identify correctly)

## Autonomous Debugging Workflow

When tests fail:

1. Check which tests failed in the output
2. Read the specific test file to understand what it's testing
3. Trace the error through the call stack shown in output
4. Make fixes to the relevant code
5. Run `/qa` again to verify the fix
6. Repeat until all tests pass (or expected failures)

## Example Output

```
=== QA Results ===
Total:   16
Passed:  11
Failed:  5
Errors:  0
Time:    53423ms
```

## When to Use

- After making any code changes to processing logic
- Before committing changes
- When debugging failing functionality
- To verify fixes work correctly
