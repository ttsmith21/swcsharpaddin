<#
.SYNOPSIS
    Compares QA test results against manifest-v2.json with field-level breakdown.

.DESCRIPTION
    Reads QA results and compares against manifest-v2.json using three-tier priority:
    csharpExpected > vbaBaseline > top-level fields.
    Reports per-part field-level status (MATCH, TOLERANCE, NOT_IMPLEMENTED, FAIL).

.PARAMETER ResultsPath
    Path to QA results JSON file. Auto-detects latest Run_* folder.

.PARAMETER ManifestPath
    Path to manifest-v2.json. Default: tests/GoldStandard_Baseline/manifest-v2.json

.PARAMETER Tolerance
    Numeric tolerance for comparisons. Default 0.05 (5%)
#>

param(
    [string]$ResultsPath = "",
    [string]$ManifestPath = "",
    [double]$Tolerance = 0.05
)

$ErrorActionPreference = "Stop"
$script:RepoRoot = Split-Path -Parent $PSScriptRoot
if (-not $script:RepoRoot -or -not (Test-Path $script:RepoRoot)) {
    $script:RepoRoot = Get-Location
}

# Auto-detect results
if (-not $ResultsPath) {
    # Try Run_Latest first
    $latestLink = Join-Path $script:RepoRoot "tests\Run_Latest\results.json"
    if (Test-Path $latestLink) {
        $ResultsPath = $latestLink
    } else {
        # Find most recent Run_* folder
        $runsDir = Join-Path $script:RepoRoot "tests"
        $latestRun = Get-ChildItem $runsDir -Directory -Filter "Run_*" -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending | Select-Object -First 1
        if ($latestRun) {
            $ResultsPath = Join-Path $latestRun.FullName "results.json"
        }
    }
}

# Default manifest-v2
if (-not $ManifestPath) {
    $ManifestPath = Join-Path $script:RepoRoot "tests\GoldStandard_Baseline\manifest-v2.json"
    if (-not (Test-Path $ManifestPath)) {
        # Fall back to v1
        $ManifestPath = Join-Path $script:RepoRoot "tests\GoldStandard_Baseline\manifest.json"
    }
}

function Write-Status($msg, $color = "White") { Write-Host $msg -ForegroundColor $color }

function Compare-Numeric($actual, $expected, $tolerance) {
    if ($null -eq $actual -or $null -eq $expected) { return $false }
    $a = [double]$actual; $e = [double]$expected
    if ($e -eq 0) { return [Math]::Abs($a) -lt 0.001 }
    $diff = [Math]::Abs($a - $e)
    $threshold = [Math]::Max([Math]::Abs($e * $tolerance), 0.001)
    return $diff -le $threshold
}

function Get-ExpectedValue($fileEntry, $fieldName) {
    # Three-tier priority: csharpExpected > vbaBaseline > top-level
    $csharp = $fileEntry.csharpExpected
    if ($csharp -and ($csharp.PSObject.Properties.Name -contains $fieldName)) {
        $v = $csharp.$fieldName
        if ($null -ne $v -and $v -ne "") { return $v }
    }
    $vba = $fileEntry.vbaBaseline
    if ($vba -and ($vba.PSObject.Properties.Name -contains $fieldName)) {
        return $vba.$fieldName
    }
    return $null
}

function Get-KnownDeviation($fileEntry, $fieldName) {
    $kd = $fileEntry.knownDeviations
    if ($kd -and ($kd.PSObject.Properties.Name -contains $fieldName)) {
        return $kd.$fieldName
    }
    return $null
}

function Compare-Field($result, $resultFieldName, $fileEntry, $manifestFieldName, $isNumeric, $tolerance) {
    $expected = Get-ExpectedValue $fileEntry $manifestFieldName
    if ($null -eq $expected) { return $null } # No expected value, skip

    $actual = $null
    if ($result.PSObject.Properties.Name -contains $resultFieldName) {
        $actual = $result.$resultFieldName
    }

    $deviation = Get-KnownDeviation $fileEntry $manifestFieldName
    $status = $deviation.status

    # Check if actual matches expected
    $matched = $false
    $toleranceMatch = $false

    if ($isNumeric) {
        if ($null -ne $actual -and $actual -ne "" -and $actual -ne 0) {
            $matched = (Compare-Numeric $actual $expected $tolerance)
            if (-not $matched) {
                $toleranceMatch = (Compare-Numeric $actual $expected ($tolerance * 2))
            }
        }
    } else {
        if ($null -ne $actual -and $actual -ne "") {
            $matched = ($actual.ToString().Trim() -ieq $expected.ToString().Trim())
        }
    }

    if ($matched) {
        return @{ Field = $manifestFieldName; Status = "MATCH"; Expected = $expected; Actual = $actual }
    }

    if ($toleranceMatch) {
        return @{ Field = $manifestFieldName; Status = "TOLERANCE"; Expected = $expected; Actual = $actual }
    }

    # Not matched - check known deviations
    if ($deviation) {
        if ($status -eq "NOT_IMPLEMENTED") {
            return @{ Field = $manifestFieldName; Status = "NOT_IMPL"; Expected = $expected; Actual = $actual; Note = $deviation.reason }
        } elseif ($status -eq "INTENTIONAL") {
            return @{ Field = $manifestFieldName; Status = "INTENTIONAL"; Expected = $expected; Actual = $actual; Note = $deviation.reason }
        } elseif ($status -eq "BUG") {
            return @{ Field = $manifestFieldName; Status = "BUG"; Expected = $expected; Actual = $actual; Note = $deviation.reason }
        }
    }

    # No deviation documented - this is a FAIL
    if ($null -eq $actual -or $actual -eq "" -or $actual -eq 0) {
        return @{ Field = $manifestFieldName; Status = "MISSING"; Expected = $expected; Actual = "(empty)" }
    }
    return @{ Field = $manifestFieldName; Status = "FAIL"; Expected = $expected; Actual = $actual }
}

# ============ MAIN ============

Write-Status ""
Write-Status "============================================" "Cyan"
Write-Status " GOLD STANDARD FIELD-LEVEL COMPARISON" "Cyan"
Write-Status "============================================" "Cyan"
Write-Status ""

if (-not $ResultsPath -or -not (Test-Path $ResultsPath)) {
    Write-Status "ERROR: Results file not found: $ResultsPath" "Red"
    Write-Status "Run QA tests first." "Yellow"
    exit 1
}
if (-not (Test-Path $ManifestPath)) {
    Write-Status "ERROR: Manifest not found: $ManifestPath" "Red"
    exit 1
}

Write-Status "Results:  $ResultsPath"
Write-Status "Manifest: $ManifestPath"
Write-Status ""

try {
    $results = Get-Content $ResultsPath -Raw | ConvertFrom-Json
    $manifest = Get-Content $ManifestPath -Raw | ConvertFrom-Json
} catch {
    Write-Status "ERROR: Failed to parse JSON: $_" "Red"
    exit 1
}

# Build lookup
$expectedByFile = @{}
foreach ($prop in $manifest.files.PSObject.Properties) {
    if ($prop.Name -notlike "_comment*") {
        $expectedByFile[$prop.Name.ToLower()] = $prop.Value
    }
}

# Field mappings: [resultField, manifestField, isNumeric]
$fieldMappings = @(
    @("Classification", "expectedClassification", $false),
    @("Thickness_in", "expectedThickness_in", $true),
    @("BendCount", "expectedBendCount", $true),
    @("OptiMaterial", "optiMaterial", $false),
    @("Description", "description", $false),
    @("Mass_lb", "rawWeight", $true),
    @("F115_Setup", "routing.N120.setup", $true),
    @("F115_Run", "routing.N120.run", $true),
    @("F140_Setup", "routing.N140.setup", $true),
    @("F140_Run", "routing.N140.run", $true),
    # F210 (debur) omitted — no baseline routing data for debur in manifest
    @("F220_Setup", "routing.N220.setup", $true),
    @("F220_Run", "routing.N220.run", $true),
    @("F325_Setup", "routing.F325.setup", $true),
    @("F325_Run", "routing.F325.run", $true)
)

# Totals
$totalMatch = 0; $totalTol = 0; $totalNotImpl = 0; $totalFail = 0; $totalMissing = 0
$totalIntentional = 0; $totalBug = 0; $totalFields = 0
$partPassCount = 0; $partFailCount = 0; $skipCount = 0

foreach ($result in $results.Results) {
    $fileName = $result.FileName.ToLower()

    if (-not $expectedByFile.ContainsKey($fileName)) {
        $skipCount++
        continue
    }

    $fileEntry = $expectedByFile[$fileName]
    $comparisons = @()

    # Check pass/fail status first
    $actualPassed = ($result.Status -eq "Success")
    $expectedPassed = if ($null -ne $fileEntry.shouldPass) { $fileEntry.shouldPass } else { $true }

    if ($actualPassed -ne $expectedPassed) {
        if ($expectedPassed) {
            $comparisons += @{ Field = "Status"; Status = "FAIL"; Expected = "Pass"; Actual = $result.Status }
        } else {
            $comparisons += @{ Field = "Status"; Status = "FAIL"; Expected = "Fail"; Actual = $result.Status }
        }
    } else {
        $comparisons += @{ Field = "Status"; Status = "MATCH"; Expected = ""; Actual = $result.Status }
    }

    # Only compare detailed fields if both sides agree the part should pass
    if ($actualPassed -and $expectedPassed) {
        # Basic fields
        foreach ($mapping in $fieldMappings) {
            $resultField = $mapping[0]
            $manifestField = $mapping[1]
            $isNumeric = $mapping[2]

            # Handle routing fields (nested in vbaBaseline.routing)
            if ($manifestField -like "routing.*") {
                $parts = $manifestField -split "\."
                $wcKey = $parts[1]
                $subField = $parts[2]

                $routingData = $null
                if ($fileEntry.vbaBaseline -and $fileEntry.vbaBaseline.routing) {
                    $routing = $fileEntry.vbaBaseline.routing
                    if ($routing.PSObject.Properties.Name -contains $wcKey) {
                        $routingData = $routing.$wcKey
                    }
                }

                if ($routingData -and ($routingData.PSObject.Properties.Name -contains $subField)) {
                    $expectedVal = $routingData.$subField
                    $actualVal = $null
                    if ($result.PSObject.Properties.Name -contains $resultField) {
                        $actualVal = $result.$resultField
                    }

                    $deviation = Get-KnownDeviation $fileEntry $manifestField
                    if ($null -ne $expectedVal) {
                        if ($null -ne $actualVal -and $actualVal -ne 0 -and (Compare-Numeric $actualVal $expectedVal $Tolerance)) {
                            $comparisons += @{ Field = $manifestField; Status = "MATCH"; Expected = $expectedVal; Actual = $actualVal }
                        } elseif ($deviation -and $deviation.status -eq "NOT_IMPLEMENTED") {
                            $comparisons += @{ Field = $manifestField; Status = "NOT_IMPL"; Expected = $expectedVal; Actual = $actualVal; Note = $deviation.reason }
                        } elseif ($null -eq $actualVal -or $actualVal -eq 0) {
                            $comparisons += @{ Field = $manifestField; Status = "MISSING"; Expected = $expectedVal; Actual = "(empty)" }
                        } else {
                            $comparisons += @{ Field = $manifestField; Status = "FAIL"; Expected = $expectedVal; Actual = $actualVal }
                        }
                    }
                }
            } else {
                $comp = Compare-Field $result $resultField $fileEntry $manifestField $isNumeric $Tolerance
                if ($comp) { $comparisons += $comp }
            }
        }
    }

    # Count statuses
    $partMatch = 0; $partTol = 0; $partNotImpl = 0; $partFail = 0; $partMissing = 0
    foreach ($c in $comparisons) {
        $totalFields++
        switch ($c.Status) {
            "MATCH"       { $partMatch++; $totalMatch++ }
            "TOLERANCE"   { $partTol++; $totalTol++ }
            "NOT_IMPL"    { $partNotImpl++; $totalNotImpl++ }
            "INTENTIONAL" { $totalIntentional++ }
            "BUG"         { $totalBug++ }
            "MISSING"     { $partMissing++; $totalMissing++ }
            "FAIL"        { $partFail++; $totalFail++ }
        }
    }

    $hasFailure = ($partFail -gt 0 -or $partMissing -gt 0)
    if ($hasFailure) { $partFailCount++ } else { $partPassCount++ }

    # Display part summary
    $summaryColor = if ($hasFailure) { "Red" } elseif ($partNotImpl -gt 0) { "Yellow" } else { "Green" }
    $statusCounts = "$partMatch MATCH"
    if ($partTol -gt 0) { $statusCounts += ", $partTol TOL" }
    if ($partNotImpl -gt 0) { $statusCounts += ", $partNotImpl NOT_IMPL" }
    if ($partMissing -gt 0) { $statusCounts += ", $partMissing MISSING" }
    if ($partFail -gt 0) { $statusCounts += ", $partFail FAIL" }

    Write-Status ("  {0}: {1}" -f $result.FileName, $statusCounts) $summaryColor

    # Show non-match details
    foreach ($c in $comparisons) {
        if ($c.Status -ne "MATCH") {
            $detailColor = switch ($c.Status) {
                "TOLERANCE"   { "DarkYellow" }
                "NOT_IMPL"    { "DarkGray" }
                "INTENTIONAL" { "DarkGray" }
                "MISSING"     { "Yellow" }
                "FAIL"        { "Red" }
                default       { "White" }
            }
            $note = if ($c.Note) { " ($($c.Note))" } else { "" }
            Write-Status ("    [{0,-9}] {1}: expected={2}, actual={3}{4}" -f $c.Status, $c.Field, $c.Expected, $c.Actual, $note) $detailColor
        }
    }
}

# Grand summary
Write-Status ""
Write-Status "============================================" "Cyan"
Write-Status " SUMMARY" "Cyan"
Write-Status "============================================" "Cyan"
Write-Status ""
Write-Status "  Parts passing:    $partPassCount" "Green"
Write-Status "  Parts with issues: $partFailCount" $(if ($partFailCount -gt 0) { "Red" } else { "Green" })
Write-Status "  Parts skipped:    $skipCount (not in manifest)" "DarkGray"
Write-Status ""
Write-Status "  Total fields:     $totalFields"
Write-Status "  Match:            $totalMatch" "Green"
Write-Status "  Tolerance:        $totalTol" $(if ($totalTol -gt 0) { "DarkYellow" } else { "Green" })
Write-Status "  Not Implemented:  $totalNotImpl" $(if ($totalNotImpl -gt 0) { "Yellow" } else { "Green" })
Write-Status "  Missing:          $totalMissing" $(if ($totalMissing -gt 0) { "Yellow" } else { "Green" })
Write-Status "  Fail:             $totalFail" $(if ($totalFail -gt 0) { "Red" } else { "Green" })
Write-Status ""

$pct = if ($totalFields -gt 0) { [Math]::Round(($totalMatch + $totalTol) / $totalFields * 100, 1) } else { 0 }
Write-Status "  Coverage: $pct% ($($totalMatch + $totalTol) / $totalFields fields match)" $(if ($pct -ge 80) { "Green" } elseif ($pct -ge 50) { "Yellow" } else { "Red" })
Write-Status ""

# Write report file
$reportPath = Join-Path (Split-Path $ResultsPath) "comparison_report.txt"
$report = @()
$report += "Gold Standard Comparison Report"
$report += "==============================="
$report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$report += "Results:   $ResultsPath"
$report += "Manifest:  $ManifestPath"
$report += ""
$report += "Summary: $totalMatch match, $totalTol tolerance, $totalNotImpl not-impl, $totalMissing missing, $totalFail fail"
$report += "Coverage: $pct% ($($totalMatch + $totalTol) / $totalFields)"
$report | Out-File $reportPath -Encoding UTF8
Write-Status "Report: $reportPath"

# Exit code: 0=pass (no FAILs), 1=failures
if ($totalFail -gt 0) { exit 1 } else { exit 0 }
