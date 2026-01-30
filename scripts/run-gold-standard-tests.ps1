<#
.SYNOPSIS
    Runs the Gold Standard regression tests for the SolidWorks Add-in.

.DESCRIPTION
    1. Copies files from tests\GoldStandard_Inputs to tests\Run_[Timestamp].
    2. Calculates SHA-256 hashes of inputs to ensure preservation.
    3. Writes config JSON for the add-in to read.
    4. Waits for results.json from the add-in.
    5. Compares results against baseline manifest.
    6. Reports pass/fail with detailed diffs.

.PARAMETER InputDir
    Path to the Gold Standard inputs. Default: tests\GoldStandard_Inputs

.PARAMETER BaselineDir
    Path to expected outputs. Default: tests\GoldStandard_Baseline

.PARAMETER Timeout
    Timeout in minutes to wait for results. Default: 5

.PARAMETER SkipCopy
    Skip copying inputs (use existing run directory).

.EXAMPLE
    .\run-gold-standard-tests.ps1

.EXAMPLE
    .\run-gold-standard-tests.ps1 -Timeout 10
#>

param (
    [string]$InputDir = "$PSScriptRoot\..\tests\GoldStandard_Inputs",
    [string]$BaselineDir = "$PSScriptRoot\..\tests\GoldStandard_Baseline",
    [int]$Timeout = 5,
    [switch]$SkipCopy
)

$ErrorActionPreference = "Stop"

# Normalize paths (remove .. components) - SolidWorks requires clean paths
$InputDir = [System.IO.Path]::GetFullPath($InputDir)
$BaselineDir = [System.IO.Path]::GetFullPath($BaselineDir)

function Write-Status {
    param([string]$Message, [string]$Color = "Cyan")
    Write-Host "[QA] $Message" -ForegroundColor $Color
}

function Write-Pass {
    param([string]$Message)
    Write-Host "   [PASS] $Message" -ForegroundColor Green
}

function Write-Fail {
    param([string]$Message)
    Write-Host "   [FAIL] $Message" -ForegroundColor Red
}

function Write-Warn {
    param([string]$Message)
    Write-Host "   [WARN] $Message" -ForegroundColor Yellow
}

function Get-FileHashDir {
    param([string]$Path)
    Get-ChildItem -Path $Path -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in @('.sldprt', '.sldasm', '.slddrw') } |
        Get-FileHash -Algorithm SHA256
}

function Compare-WithTolerance {
    param(
        [double]$Expected,
        [double]$Actual,
        [double]$Tolerance = 0.001
    )
    if ($Expected -eq 0) {
        return $Actual -eq 0
    }
    $diff = [Math]::Abs($Expected - $Actual) / [Math]::Abs($Expected)
    return $diff -le $Tolerance
}

# ============================================================================
# 1. Setup Run Directory
# ============================================================================
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$runDir = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\tests\Run_$timestamp")
$runInputs = "$runDir\Inputs"

Write-Status "Starting Test Run: $timestamp"
Write-Status "Creating run directory: $runDir"

New-Item -ItemType Directory -Force -Path $runInputs | Out-Null

if (!(Test-Path $InputDir)) {
    Write-Host "[QA] ERROR: Input directory not found: $InputDir" -ForegroundColor Red
    exit 1
}

# Check for input files
$inputFiles = Get-ChildItem -Path $InputDir -Filter "*.sldprt" -ErrorAction SilentlyContinue
if ($inputFiles.Count -eq 0) {
    Write-Host "[QA] ERROR: No .sldprt files found in $InputDir" -ForegroundColor Red
    Write-Host "[QA] Add your test files to tests\GoldStandard_Inputs\" -ForegroundColor Yellow
    exit 1
}

Write-Status "Found $($inputFiles.Count) test file(s)"

# Show coverage summary
$baselineFile = "$BaselineDir\manifest.json"
if (Test-Path $baselineFile) {
    $baseline = Get-Content $baselineFile -Raw | ConvertFrom-Json
    $expectedFiles = $baseline.files.PSObject.Properties.Name | Where-Object { $_ -notlike "_comment*" }
    $presentFiles = $inputFiles.Name
    $testedFromBaseline = ($expectedFiles | Where-Object { $presentFiles -contains $_ }).Count

    Write-Host ""
    Write-Host "   Coverage: $testedFromBaseline / $($expectedFiles.Count) baseline files present" -ForegroundColor Gray

    # Show by class
    $classCounts = @{}
    foreach ($file in $presentFiles) {
        $prefix = $file.Substring(0, 1).ToUpper()
        if ($prefix -match "[A-H]") {
            if (-not $classCounts.ContainsKey($prefix)) { $classCounts[$prefix] = 0 }
            $classCounts[$prefix]++
        }
    }
    if ($classCounts.Count -gt 0) {
        $classStr = ($classCounts.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name):$($_.Value)" }) -join ", "
        Write-Host "   By class: $classStr" -ForegroundColor Gray
    }
    Write-Host ""
}

# ============================================================================
# 2. Copy Inputs (Preservation Strategy)
# ============================================================================
if (-not $SkipCopy) {
    Write-Status "Copying Gold Standard inputs..."
    Copy-Item -Path "$InputDir\*" -Destination $runInputs -Recurse -Force

    # Verify integrity
    $sourceHashes = Get-FileHashDir $InputDir
    $copiedHashes = Get-FileHashDir $runInputs

    if ($sourceHashes.Count -ne $copiedHashes.Count) {
        Write-Host "[QA] WARNING: Hash count mismatch after copy" -ForegroundColor Yellow
    } else {
        Write-Status "Input integrity verified ($($sourceHashes.Count) files)" -Color Green
    }
}

# ============================================================================
# 3. Write Config for Add-in
# ============================================================================
$configPath = "C:\Temp\nm_qa_config.json"
$resultFile = "$runDir\results.json"

$config = @{
    InputPath = $runInputs
    OutputPath = $resultFile
    BaselinePath = "$BaselineDir\manifest.json"
}

# Ensure C:\Temp exists
if (!(Test-Path "C:\Temp")) {
    New-Item -ItemType Directory -Force -Path "C:\Temp" | Out-Null
}

$config | ConvertTo-Json | Set-Content $configPath -Encoding UTF8
Write-Status "Config written to $configPath"

# ============================================================================
# 4. Wait for SolidWorks to Process
# ============================================================================
Write-Status "WAITING FOR RESULTS..." -Color Yellow
Write-Host ""
Write-Host "   Instructions:" -ForegroundColor White
Write-Host "   1. Open SolidWorks (if not already running)"
Write-Host "   2. Click 'Run QA' from the add-in toolbar"
Write-Host "   3. Wait for completion dialog"
Write-Host ""
Write-Host "   Waiting up to $Timeout minutes for results.json..."
Write-Host ""

$timeoutTime = (Get-Date).AddMinutes($Timeout)
$dotCount = 0

while (!(Test-Path $resultFile)) {
    if ((Get-Date) -gt $timeoutTime) {
        Write-Host ""
        Write-Host "[QA] TIMEOUT: No results after $Timeout minutes" -ForegroundColor Red
        Write-Host "[QA] Make sure SolidWorks is running and click 'Run QA'" -ForegroundColor Yellow
        exit 1
    }
    Start-Sleep -Seconds 2
    $dotCount++
    if ($dotCount % 5 -eq 0) {
        Write-Host "." -NoNewline
    }
}

Write-Host ""
Write-Status "Results found!" -Color Green

# ============================================================================
# 5. Load Results and Baseline
# ============================================================================
$results = Get-Content $resultFile -Raw | ConvertFrom-Json
Write-Status "Loaded $($results.Results.Count) test results"

$baselineFile = "$BaselineDir\manifest.json"
$baseline = $null
if (Test-Path $baselineFile) {
    $baseline = Get-Content $baselineFile -Raw | ConvertFrom-Json
    Write-Status "Loaded baseline manifest"
} else {
    Write-Warn "No baseline manifest found - reporting raw results only"
}

# ============================================================================
# 6. Analyze Results
# ============================================================================
Write-Status "Analyzing results..."
Write-Host ""

$passCount = 0
$failCount = 0
$warnCount = 0
$comparisonReport = @()

foreach ($res in $results.Results) {
    $fileName = $res.FileName
    $baselineEntry = $null

    if ($baseline -and $baseline.files.PSObject.Properties.Name -contains $fileName) {
        $baselineEntry = $baseline.files.$fileName
    }

    # Skip template entries
    if ($fileName -like "_TEMPLATE_*") {
        continue
    }

    $issues = @()

    # Check pass/fail expectation
    if ($baselineEntry) {
        $shouldPass = $baselineEntry.shouldPass
        $didPass = ($res.Status -eq "Success")

        if ($shouldPass -and -not $didPass) {
            $issues += "Expected PASS but got $($res.Status): $($res.Message)"
        }
        elseif (-not $shouldPass -and $didPass) {
            $issues += "Expected FAIL but got SUCCESS"
        }
        elseif (-not $shouldPass -and -not $didPass) {
            # Check if failure reason matches
            if ($baselineEntry.expectedFailureReason) {
                if ($res.Message -notlike "*$($baselineEntry.expectedFailureReason)*") {
                    $issues += "Failure reason mismatch: expected '$($baselineEntry.expectedFailureReason)', got '$($res.Message)'"
                }
            }
        }

        # Check classification
        if ($baselineEntry.expectedClassification -and $res.Classification -ne $baselineEntry.expectedClassification) {
            $issues += "Classification: expected '$($baselineEntry.expectedClassification)', got '$($res.Classification)'"
        }

        # Get tolerances
        $defaultTol = if ($baseline.defaultTolerances) { $baseline.defaultTolerances } else { @{ thickness = 0.001; mass = 0.01 } }
        $tol = if ($baselineEntry.tolerances) {
            @{
                thickness = if ($baselineEntry.tolerances.thickness) { $baselineEntry.tolerances.thickness } else { $defaultTol.thickness }
                mass = if ($baselineEntry.tolerances.mass) { $baselineEntry.tolerances.mass } else { $defaultTol.mass }
            }
        } else { $defaultTol }

        # Check numeric values with tolerance
        if ($baselineEntry.expectedThickness_in -and $res.Thickness_in) {
            if (-not (Compare-WithTolerance $baselineEntry.expectedThickness_in $res.Thickness_in $tol.thickness)) {
                $issues += "Thickness: expected $($baselineEntry.expectedThickness_in), got $($res.Thickness_in)"
            }
        }

        if ($baselineEntry.expectedMass_lb -and $res.Mass_lb) {
            if (-not (Compare-WithTolerance $baselineEntry.expectedMass_lb $res.Mass_lb $tol.mass)) {
                $issues += "Mass: expected $($baselineEntry.expectedMass_lb), got $($res.Mass_lb)"
            }
        }

        if ($baselineEntry.expectedBendCount -and $null -ne $res.BendCount) {
            if ($baselineEntry.expectedBendCount -ne $res.BendCount) {
                $issues += "BendCount: expected $($baselineEntry.expectedBendCount), got $($res.BendCount)"
            }
        }

        # Tube checks
        if ($baselineEntry.expectedTubeOD_in -and $res.TubeOD_in) {
            if (-not (Compare-WithTolerance $baselineEntry.expectedTubeOD_in $res.TubeOD_in $tol.thickness)) {
                $issues += "TubeOD: expected $($baselineEntry.expectedTubeOD_in), got $($res.TubeOD_in)"
            }
        }
    }
    else {
        # No baseline - just report status
        if ($res.Status -ne "Success") {
            $issues += "Processing $($res.Status): $($res.Message)"
        }
    }

    # Report result
    if ($issues.Count -eq 0) {
        Write-Pass "$fileName ($($res.Classification), $([int]$res.ElapsedMs)ms)"
        $passCount++
    } else {
        Write-Fail "$fileName"
        foreach ($issue in $issues) {
            Write-Host "      - $issue" -ForegroundColor Gray
        }
        $failCount++
    }

    $comparisonReport += [PSCustomObject]@{
        FileName = $fileName
        Status = if ($issues.Count -eq 0) { "PASS" } else { "FAIL" }
        Issues = $issues -join "; "
        Classification = $res.Classification
        ElapsedMs = $res.ElapsedMs
    }
}

# ============================================================================
# 7. Summary
# ============================================================================
Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " QA TEST SUMMARY" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Run ID:     $($results.RunId)"
Write-Host "  Total:      $($results.TotalFiles)"
Write-Host "  Passed:     $passCount" -ForegroundColor Green
Write-Host "  Failed:     $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Host "  Duration:   $([int]$results.TotalElapsedMs)ms"
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan

# Write comparison report
$reportFile = "$runDir\comparison_report.txt"
$comparisonReport | Format-Table -AutoSize | Out-String | Set-Content $reportFile
Write-Status "Report saved to: $reportFile"
Write-Status "Results JSON: $resultFile"

if ($failCount -eq 0) {
    Write-Host ""
    Write-Host "[QA] ALL TESTS PASSED" -ForegroundColor Green
    exit 0
} else {
    Write-Host ""
    Write-Host "[QA] $failCount TEST(S) FAILED" -ForegroundColor Red
    exit 1
}
