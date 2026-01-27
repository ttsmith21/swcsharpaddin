<#
.SYNOPSIS
    Builds the project and runs tests, providing a summary.

.DESCRIPTION
    1. Optionally cleans previous build artifacts
    2. Builds swcsharpaddin using Visual Studio MSBuild (required for COM interop)
    3. Runs unit tests in NM.Core.Tests using dotnet test
    4. Summarizes results (pass/fail, error count)

.PARAMETER SkipTests
    If specified, only builds without running tests.

.PARAMETER SkipClean
    If specified, skips the clean step for faster incremental builds.

.PARAMETER Verbose
    If specified, shows full build/test output instead of summary.
#>
param(
    [switch]$SkipTests,
    [switch]$SkipClean,
    [switch]$Verbose
)

$ErrorActionPreference = "Continue"
$projectRoot = Split-Path -Parent $PSScriptRoot
$csprojPath = Join-Path $projectRoot "swcsharpaddin.csproj"
$testProjectPath = Join-Path $projectRoot "src\NM.Core.Tests"

Write-Host "=== build-and-test ===" -ForegroundColor Cyan
Write-Host "Project root: $projectRoot"
Write-Host ""

# Track results
$buildSuccess = $false
$testSuccess = $false

# Find MSBuild (VS MSBuild required for COM interop projects)
$msbuildPaths = @(
    "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\18\Insiders\MSBuild\Current\Bin\amd64\MSBuild.exe"
)

$msbuild = $null
foreach ($path in $msbuildPaths) {
    if (Test-Path $path) {
        $msbuild = $path
        break
    }
}

if (-not $msbuild) {
    Write-Host "ERROR: Could not find Visual Studio MSBuild" -ForegroundColor Red
    Write-Host "  dotnet build cannot be used with RegisterForComInterop projects" -ForegroundColor Yellow
    exit 1
}

Write-Host "Using MSBuild: $msbuild"
Write-Host ""

# === CLEAN ===
if (-not $SkipClean) {
    Write-Host "--- Cleaning ---" -ForegroundColor Yellow
    $binPath = Join-Path $projectRoot "bin"
    $objPath = Join-Path $projectRoot "obj"
    if (Test-Path $binPath) {
        Remove-Item -Path $binPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path $objPath) {
        Remove-Item -Path $objPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    Write-Host "Clean complete (bin/obj deleted)"
    Write-Host ""
}

# === BUILD ===
Write-Host "--- Building ---" -ForegroundColor Yellow

$buildOutput = & $msbuild $csprojPath "-p:Configuration=Debug" "-v:normal" "-nologo" "-m" 2>&1 | Out-String
$buildExitCode = $LASTEXITCODE

# Parse build output for errors and warnings
$buildLines = $buildOutput -split "`n"
$buildErrors = @()
$buildWarnings = @()

foreach ($line in $buildLines) {
    if ($line -match ": error [A-Z]+\d+:") {
        # Skip COM registration errors (MSB3216, MSB3392) - these require admin
        if ($line -notmatch "MSB3216|MSB3392") {
            $buildErrors += $line.Trim()
        }
    }
    elseif ($line -match ": warning [A-Z]+\d+:") {
        $buildWarnings += $line.Trim()
    }
}

# Check if DLL was created (build success even if COM registration failed)
$dllPath = Join-Path $projectRoot "bin\Debug\swcsharpaddin.dll"
$dllCreated = Test-Path $dllPath

if ($dllCreated -and $buildErrors.Count -eq 0) {
    $buildSuccess = $true
    Write-Host "BUILD: SUCCESS" -ForegroundColor Green
    if ($buildWarnings.Count -gt 0) {
        Write-Host "  Warnings: $($buildWarnings.Count)" -ForegroundColor Yellow
        if ($buildWarnings.Count -le 10 -or $Verbose) {
            foreach ($warn in $buildWarnings) {
                Write-Host "    $warn" -ForegroundColor Yellow
            }
        } else {
            for ($i = 0; $i -lt 10; $i++) {
                Write-Host "    $($buildWarnings[$i])" -ForegroundColor Yellow
            }
            Write-Host "    ... and $($buildWarnings.Count - 10) more warnings" -ForegroundColor Yellow
        }
    }
    # Check for COM registration issues (warning, not error)
    if ($buildOutput -match "MSB3216|MSB3392") {
        Write-Host "  Note: COM registration skipped (requires admin)" -ForegroundColor Gray
    }
} else {
    Write-Host "BUILD: FAILED" -ForegroundColor Red
    if ($buildErrors.Count -gt 0) {
        Write-Host "  Errors: $($buildErrors.Count)" -ForegroundColor Red
        $showCount = [Math]::Min($buildErrors.Count, 15)
        for ($i = 0; $i -lt $showCount; $i++) {
            Write-Host "    $($buildErrors[$i])" -ForegroundColor Red
        }
        if ($buildErrors.Count -gt 15) {
            Write-Host "    ... and $($buildErrors.Count - 15) more errors" -ForegroundColor Red
        }
    } else {
        Write-Host "  DLL not created" -ForegroundColor Red
    }
}

if ($Verbose) {
    Write-Host ""
    Write-Host "--- Full Build Output ---" -ForegroundColor Gray
    Write-Host $buildOutput
}

Write-Host ""

# === TESTS ===
if (-not $SkipTests -and $buildSuccess) {
    Write-Host "--- Testing ---" -ForegroundColor Yellow

    if (-not (Test-Path $testProjectPath)) {
        Write-Host "TEST: SKIPPED (test project not found)" -ForegroundColor Yellow
    } else {
        # Build and run tests (test project references pre-built DLL, not ProjectReference)
        $testOutput = & dotnet test $testProjectPath --verbosity minimal 2>&1 | Out-String

        if ($testOutput -match "Passed!\s+- Failed:\s+(\d+), Passed:\s+(\d+), Skipped:\s+(\d+), Total:\s+(\d+)") {
            $failed = [int]$Matches[1]
            $passed = [int]$Matches[2]
            $skipped = [int]$Matches[3]
            $total = [int]$Matches[4]

            if ($failed -eq 0) {
                $testSuccess = $true
                Write-Host "TEST: PASSED ($passed/$total)" -ForegroundColor Green
                if ($skipped -gt 0) {
                    Write-Host "  Skipped: $skipped" -ForegroundColor Yellow
                }
            } else {
                Write-Host "TEST: FAILED ($failed failed, $passed passed)" -ForegroundColor Red
            }
        }
        elseif ($testOutput -match "Failed!\s+- Failed:\s+(\d+), Passed:\s+(\d+)") {
            $failed = [int]$Matches[1]
            $passed = [int]$Matches[2]
            Write-Host "TEST: FAILED ($failed failed, $passed passed)" -ForegroundColor Red
        }
        elseif ($testOutput -match "error") {
            Write-Host "TEST: ERROR (build or runtime error)" -ForegroundColor Red
            # Try building tests first
            Write-Host "  Attempting to build test project..." -ForegroundColor Yellow
            $testOutput = & dotnet test $testProjectPath --verbosity minimal 2>&1 | Out-String
            if ($testOutput -match "Passed!\s+- Failed:\s+(\d+), Passed:\s+(\d+), Skipped:\s+(\d+), Total:\s+(\d+)") {
                $failed = [int]$Matches[1]
                $passed = [int]$Matches[2]
                $total = [int]$Matches[4]

                if ($failed -eq 0) {
                    $testSuccess = $true
                    Write-Host "TEST: PASSED ($passed/$total)" -ForegroundColor Green
                } else {
                    Write-Host "TEST: FAILED ($failed failed, $passed passed)" -ForegroundColor Red
                }
            }
        }
        else {
            # Try building tests first if --no-build failed
            $testOutput = & dotnet test $testProjectPath --verbosity minimal 2>&1 | Out-String
            if ($testOutput -match "Passed!\s+- Failed:\s+(\d+), Passed:\s+(\d+), Skipped:\s+(\d+), Total:\s+(\d+)") {
                $failed = [int]$Matches[1]
                $passed = [int]$Matches[2]
                $total = [int]$Matches[4]

                if ($failed -eq 0) {
                    $testSuccess = $true
                    Write-Host "TEST: PASSED ($passed/$total)" -ForegroundColor Green
                } else {
                    Write-Host "TEST: FAILED ($failed failed, $passed passed)" -ForegroundColor Red
                }
            } else {
                Write-Host "TEST: UNKNOWN (could not parse output)" -ForegroundColor Yellow
            }
        }

        if ($Verbose) {
            Write-Host ""
            Write-Host "--- Full Test Output ---" -ForegroundColor Gray
            Write-Host $testOutput
        }
    }
} elseif (-not $buildSuccess) {
    Write-Host "--- Testing ---" -ForegroundColor Yellow
    Write-Host "TEST: SKIPPED (build failed)" -ForegroundColor Yellow
}

Write-Host ""

# === SUMMARY ===
Write-Host "=== Summary ===" -ForegroundColor Cyan
$overallSuccess = $buildSuccess -and ($SkipTests -or $testSuccess)

if ($overallSuccess) {
    Write-Host "OVERALL: SUCCESS" -ForegroundColor Green
    if ($buildWarnings.Count -gt 0) {
        Write-Host "  ($($buildWarnings.Count) warnings)" -ForegroundColor Yellow
    }
    exit 0
} else {
    Write-Host "OVERALL: FAILED" -ForegroundColor Red
    if (-not $buildSuccess) {
        Write-Host "  Build failed"
    }
    if (-not $SkipTests -and -not $testSuccess -and $buildSuccess) {
        Write-Host "  Tests failed"
    }
    exit 1
}
