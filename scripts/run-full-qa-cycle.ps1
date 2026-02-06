<#
.SYNOPSIS
    Single-command orchestrator for the full QA feedback loop.
.DESCRIPTION
    Builds the project, runs headless QA tests against gold standard inputs,
    and compares results against manifest-v2.json with field-level breakdown.
.PARAMETER SkipBuild
    Skip the build step (use existing DLL).
.PARAMETER SkipQA
    Skip the QA run (use existing results.json).
.PARAMETER ManifestV2
    Path to manifest-v2.json. Default: tests/GoldStandard_Baseline/manifest-v2.json
#>
param(
    [switch]$SkipBuild,
    [switch]$SkipQA,
    [string]$ManifestV2 = "$PSScriptRoot\..\tests\GoldStandard_Baseline\manifest-v2.json"
)

$ErrorActionPreference = "Stop"
$projectRoot = (Resolve-Path "$PSScriptRoot\..").Path

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " GOLD STANDARD QA CYCLE" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$overallStart = Get-Date

# --- Step 1: Kill SolidWorks if running ---
Write-Host "[1/7] Checking for SolidWorks..." -ForegroundColor Yellow
$swProcess = Get-Process -Name SLDWORKS -ErrorAction SilentlyContinue
if ($swProcess) {
    Write-Host "  SolidWorks is running. Stopping..." -ForegroundColor Red
    $swProcess | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    Write-Host "  SolidWorks stopped." -ForegroundColor Green
} else {
    Write-Host "  SolidWorks not running." -ForegroundColor Green
}

# --- Step 2: Build ---
if (-not $SkipBuild) {
    Write-Host ""
    Write-Host "[2/7] Building..." -ForegroundColor Yellow

    # Build main project
    & "$PSScriptRoot\build-and-test.ps1" -SkipClean -SkipTests
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  BUILD FAILED" -ForegroundColor Red
        exit 1
    }

    # Build BatchRunner
    $msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe"
    if (Test-Path $msbuild) {
        & $msbuild "$projectRoot\src\NM.BatchRunner\NM.BatchRunner.csproj" /p:Configuration=Debug /v:minimal
        if ($LASTEXITCODE -ne 0) {
            Write-Host "  BatchRunner build FAILED" -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host "  WARNING: MSBuild not found, skipping BatchRunner build" -ForegroundColor Yellow
    }

    # Verify DLL timestamp
    $dll = Get-Item "$projectRoot\bin\Debug\swcsharpaddin.dll" -ErrorAction SilentlyContinue
    if ($dll) {
        $age = (Get-Date) - $dll.LastWriteTime
        if ($age.TotalMinutes -gt 5) {
            Write-Host "  WARNING: DLL is $([int]$age.TotalMinutes) minutes old - may be stale" -ForegroundColor Yellow
        } else {
            Write-Host "  DLL built: $($dll.LastWriteTime.ToString('HH:mm:ss'))" -ForegroundColor Green
        }
    }
} else {
    Write-Host ""
    Write-Host "[2/7] Skipping build (--SkipBuild)" -ForegroundColor DarkGray
}

# --- Step 3: Run QA ---
if (-not $SkipQA) {
    Write-Host ""
    Write-Host "[3/7] Running QA tests..." -ForegroundColor Yellow
    $batchRunner = "$projectRoot\src\NM.BatchRunner\bin\Debug\NM.BatchRunner.exe"
    if (-not (Test-Path $batchRunner)) {
        Write-Host "  ERROR: NM.BatchRunner.exe not found at: $batchRunner" -ForegroundColor Red
        Write-Host "  Run without -SkipBuild to build it first." -ForegroundColor Red
        exit 1
    }

    $qaStart = Get-Date
    & $batchRunner --qa
    $qaExitCode = $LASTEXITCODE
    $qaElapsed = (Get-Date) - $qaStart

    if ($qaExitCode -ne 0) {
        Write-Host "  QA run failed (exit code: $qaExitCode)" -ForegroundColor Red
    } else {
        Write-Host "  QA completed in $([int]$qaElapsed.TotalSeconds)s" -ForegroundColor Green
    }
} else {
    Write-Host ""
    Write-Host "[3/7] Skipping QA run (--SkipQA)" -ForegroundColor DarkGray
}

# --- Step 4: Find latest results ---
Write-Host ""
Write-Host "[4/7] Finding results..." -ForegroundColor Yellow
$runsDir = "$projectRoot\tests"
$latestRun = Get-ChildItem $runsDir -Directory -Filter "Run_*" -ErrorAction SilentlyContinue |
    Sort-Object Name -Descending |
    Select-Object -First 1

$resultsFile = $null
if ($latestRun) {
    $resultsFile = Join-Path $latestRun.FullName "results.json"
    if (-not (Test-Path $resultsFile)) { $resultsFile = $null }
}

# Also check Run_Latest symlink
if (-not $resultsFile) {
    $latestLink = "$runsDir\Run_Latest\results.json"
    if (Test-Path $latestLink) { $resultsFile = $latestLink }
}

if (-not $resultsFile) {
    Write-Host "  No results.json found in test runs." -ForegroundColor Red
    Write-Host "  Run QA first: .\scripts\run-full-qa-cycle.ps1" -ForegroundColor Red
    exit 1
}

Write-Host "  Results: $resultsFile" -ForegroundColor Green

# --- Step 5: Compare against manifest-v2 ---
Write-Host ""
Write-Host "[5/7] Comparing results against manifest-v2..." -ForegroundColor Yellow

if (-not (Test-Path $ManifestV2)) {
    Write-Host "  manifest-v2.json not found at: $ManifestV2" -ForegroundColor Red
    Write-Host "  Run: .\scripts\import-prn-to-manifest.ps1" -ForegroundColor Red
    exit 1
}

# Use the existing compare script if it supports --manifest-v2
$compareScript = "$PSScriptRoot\compare-qa-results.ps1"
if (Test-Path $compareScript) {
    Write-Host "  Running comparison script..." -ForegroundColor Cyan
    & $compareScript
} else {
    Write-Host "  compare-qa-results.ps1 not found, showing basic results summary" -ForegroundColor Yellow
}

# --- Step 6: Import.prn diff ---
Write-Host ""
Write-Host "[6/7] Import.prn diff..." -ForegroundColor Yellow
$diffScript = "$PSScriptRoot\diff-import-prn.ps1"
if (Test-Path $diffScript) {
    $csharpPrn = Join-Path (Split-Path $resultsFile) "Import_CSharp.prn"
    if (Test-Path $csharpPrn) {
        & $diffScript
    } else {
        Write-Host "  No Import_CSharp.prn found, skipping diff" -ForegroundColor DarkGray
    }
} else {
    Write-Host "  diff-import-prn.ps1 not found" -ForegroundColor DarkGray
}

# --- Step 7: Track coverage ---
Write-Host ""
Write-Host "[7/7] Tracking coverage..." -ForegroundColor Yellow
$trackScript = "$PSScriptRoot\track-coverage.ps1"
if (Test-Path $trackScript) {
    & $trackScript
} else {
    Write-Host "  track-coverage.ps1 not found" -ForegroundColor DarkGray
}

# Parse results.json for a quick summary
$resultsJson = Get-Content $resultsFile -Raw | ConvertFrom-Json
$totalParts = 0
$passCount = 0
$failCount = 0
$errorCount = 0

if ($resultsJson.Results) {
    foreach ($r in $resultsJson.Results) {
        $totalParts++
        switch ($r.Status) {
            "Success" { $passCount++ }
            "Failed" { $failCount++ }
            default { $errorCount++ }
        }
    }
}

$overallElapsed = (Get-Date) - $overallStart

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " QA CYCLE COMPLETE" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Total time:  $([int]$overallElapsed.TotalSeconds)s" -ForegroundColor White
Write-Host "  Parts tested: $totalParts" -ForegroundColor White
Write-Host "  Passed:       $passCount" -ForegroundColor Green
Write-Host "  Failed:       $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Host "  Errors:       $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
Write-Host ""

# Quick per-part status table
if ($resultsJson.Results) {
    Write-Host "Per-Part Results:" -ForegroundColor Cyan
    foreach ($r in $resultsJson.Results | Sort-Object FileName) {
        $statusColor = switch ($r.Status) {
            "Success" { "Green" }
            "Failed" { "Yellow" }
            default { "Red" }
        }
        $classification = if ($r.Classification) { $r.Classification } else { "-" }
        $optiMat = if ($r.OptiMaterial) { $r.OptiMaterial } else { "(none)" }
        Write-Host ("  {0,-45} {1,-8} {2,-12} {3}" -f $r.FileName, $r.Status, $classification, $optiMat) -ForegroundColor $statusColor
    }
}

Write-Host ""
if ($failCount -gt 0 -or $errorCount -gt 0) {
    Write-Host "RESULT: FAILURES DETECTED" -ForegroundColor Red
    exit 1
} else {
    Write-Host "RESULT: ALL PASSED" -ForegroundColor Green
    exit 0
}
