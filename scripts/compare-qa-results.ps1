<#
.SYNOPSIS
    Compares QA test results against the Gold Standard baseline manifest.

.DESCRIPTION
    Reads the latest QA results from tests/Run_Latest/results.json and compares
    against expected values in tests/GoldStandard_Baseline/manifest.json.
    Reports any mismatches in classification, dimensions, or pass/fail status.

.PARAMETER ResultsPath
    Path to QA results JSON file. Defaults to tests/Run_Latest/results.json

.PARAMETER ManifestPath
    Path to baseline manifest JSON file. Defaults to tests/GoldStandard_Baseline/manifest.json

.PARAMETER Tolerance
    Numeric tolerance for dimension comparisons. Default 0.01 (1%)

.EXAMPLE
    .\compare-qa-results.ps1

.EXAMPLE
    .\compare-qa-results.ps1 -ResultsPath "tests/Run_20260129/results.json"
#>

param(
    [string]$ResultsPath = "",
    [string]$ManifestPath = "",
    [double]$Tolerance = 0.01
)

$ErrorActionPreference = "Stop"
# Get repo root - script is in scripts/ folder
$script:RepoRoot = Split-Path -Parent $PSScriptRoot
if (-not $script:RepoRoot -or -not (Test-Path $script:RepoRoot)) {
    $script:RepoRoot = Get-Location
}

# Default paths relative to repo root
if (-not $ResultsPath) {
    $ResultsPath = Join-Path $script:RepoRoot "tests\Run_Latest\results.json"
}
if (-not $ManifestPath) {
    $ManifestPath = Join-Path $script:RepoRoot "tests\GoldStandard_Baseline\manifest.json"
}

function Write-Status($msg, $color = "White") {
    Write-Host $msg -ForegroundColor $color
}

function Compare-WithTolerance($actual, $expected, $tolerance) {
    if ($null -eq $actual -or $null -eq $expected) { return $false }
    $diff = [Math]::Abs($actual - $expected)
    $threshold = [Math]::Abs($expected * $tolerance)
    if ($threshold -lt 0.001) { $threshold = 0.001 }  # Minimum absolute threshold
    return $diff -le $threshold
}

# ============ MAIN ============

Write-Status "=== QA Results Comparison ===" "Cyan"
Write-Status ""

# Check files exist
if (-not (Test-Path $ResultsPath)) {
    Write-Status "ERROR: Results file not found: $ResultsPath" "Red"
    Write-Status "Run QA tests first to generate results." "Yellow"
    exit 1
}

if (-not (Test-Path $ManifestPath)) {
    Write-Status "ERROR: Manifest file not found: $ManifestPath" "Red"
    exit 1
}

Write-Status "Results:  $ResultsPath"
Write-Status "Manifest: $ManifestPath"
Write-Status "Tolerance: $($Tolerance * 100)%"
Write-Status ""

# Load JSON files
try {
    $results = Get-Content $ResultsPath -Raw | ConvertFrom-Json
    $manifest = Get-Content $ManifestPath -Raw | ConvertFrom-Json
} catch {
    Write-Status "ERROR: Failed to parse JSON: $_" "Red"
    exit 1
}

# Build lookup from manifest (normalize to lowercase for matching)
$expectedByFile = @{}
foreach ($prop in $manifest.files.PSObject.Properties) {
    if ($prop.Name -notlike "_comment*") {
        $key = $prop.Name.ToLower()
        $expectedByFile[$key] = $prop.Value
    }
}

# Compare each result
$passCount = 0
$failCount = 0
$missingCount = 0
$issues = @()

foreach ($result in $results.Results) {
    $fileName = $result.FileName.ToLower()

    if (-not $expectedByFile.ContainsKey($fileName)) {
        $missingCount++
        Write-Status "  [SKIP] $($result.FileName) - not in manifest" "DarkGray"
        continue
    }

    $expected = $expectedByFile[$fileName]
    $fileIssues = @()

    # Check pass/fail status
    $actualPassed = ($result.Status -eq "Success")
    $expectedPassed = $expected.shouldPass

    if ($actualPassed -ne $expectedPassed) {
        if ($expectedPassed) {
            $fileIssues += "Should PASS but got Status='$($result.Status)' Message='$($result.Message)'"
        } else {
            $fileIssues += "Should FAIL ('$($expected.expectedFailureReason)') but got Status='$($result.Status)'"
        }
    }

    # If both passed, check classification
    if ($actualPassed -and $expectedPassed -and $expected.expectedClassification) {
        if ($result.Classification -ne $expected.expectedClassification) {
            $fileIssues += "Classification: expected '$($expected.expectedClassification)' got '$($result.Classification)'"
        }
    }

    # Check numeric values if both passed
    if ($actualPassed -and $expectedPassed) {
        # Thickness
        if ($expected.expectedThickness_in -and $result.Thickness_in) {
            if (-not (Compare-WithTolerance $result.Thickness_in $expected.expectedThickness_in $Tolerance)) {
                $fileIssues += "Thickness: expected $($expected.expectedThickness_in) got $($result.Thickness_in)"
            }
        }

        # Tube OD
        if ($expected.expectedTubeOD_in -and $result.TubeOD_in) {
            if (-not (Compare-WithTolerance $result.TubeOD_in $expected.expectedTubeOD_in $Tolerance)) {
                $fileIssues += "TubeOD: expected $($expected.expectedTubeOD_in) got $($result.TubeOD_in)"
            }
        }

        # Tube Wall
        if ($expected.expectedTubeWall_in -and $result.TubeWall_in) {
            if (-not (Compare-WithTolerance $result.TubeWall_in $expected.expectedTubeWall_in $Tolerance)) {
                $fileIssues += "TubeWall: expected $($expected.expectedTubeWall_in) got $($result.TubeWall_in)"
            }
        }

        # Tube ID
        if ($expected.expectedTubeID_in -and $result.TubeID_in) {
            if (-not (Compare-WithTolerance $result.TubeID_in $expected.expectedTubeID_in $Tolerance)) {
                $fileIssues += "TubeID: expected $($expected.expectedTubeID_in) got $($result.TubeID_in)"
            }
        }

        # Tube Length
        if ($expected.expectedTubeLength_in -and $result.TubeLength_in) {
            if (-not (Compare-WithTolerance $result.TubeLength_in $expected.expectedTubeLength_in $Tolerance)) {
                $fileIssues += "TubeLength: expected $($expected.expectedTubeLength_in) got $($result.TubeLength_in)"
            }
        }

        # Bend count
        if ($expected.expectedBendCount -and $result.BendCount) {
            if ($result.BendCount -ne $expected.expectedBendCount) {
                $fileIssues += "BendCount: expected $($expected.expectedBendCount) got $($result.BendCount)"
            }
        }
    }

    # Report result
    if ($fileIssues.Count -eq 0) {
        $passCount++
        Write-Status "  [PASS] $($result.FileName)" "Green"
    } else {
        $failCount++
        Write-Status "  [FAIL] $($result.FileName)" "Red"
        foreach ($issue in $fileIssues) {
            Write-Status "         - $issue" "Yellow"
            $issues += @{ File = $result.FileName; Issue = $issue }
        }
    }
}

# Summary
Write-Status ""
Write-Status "=== Summary ===" "Cyan"
Write-Status "  Passed:  $passCount" "Green"
Write-Status "  Failed:  $failCount" $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Status "  Skipped: $missingCount (not in manifest)" "DarkGray"
Write-Status ""

# Write detailed report
$reportPath = Join-Path (Split-Path $ResultsPath) "comparison_report.txt"
$report = @()
$report += "QA Comparison Report"
$report += "===================="
$report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$report += "Results:   $ResultsPath"
$report += "Manifest:  $ManifestPath"
$report += ""
$report += "Summary: $passCount passed, $failCount failed, $missingCount skipped"
$report += ""

if ($issues.Count -gt 0) {
    $report += "Issues:"
    $report += "-------"
    foreach ($issue in $issues) {
        $report += "  $($issue.File): $($issue.Issue)"
    }
}

$report | Out-File $reportPath -Encoding UTF8
Write-Status "Report saved to: $reportPath"

# Exit code
if ($failCount -gt 0) {
    exit 1
} else {
    exit 0
}
