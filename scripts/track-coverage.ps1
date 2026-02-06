<#
.SYNOPSIS
    Tracks gold standard comparison coverage over time.

.DESCRIPTION
    After each QA run, parses the comparison_report.txt and appends metrics
    to tests/coverage-history.csv. Displays trend of last N runs showing
    NOT_IMPLEMENTED decreasing and MATCH increasing.

.PARAMETER ResultsDir
    Path to QA results directory. Auto-detects latest Run_*.

.PARAMETER ShowTrend
    Number of recent runs to display in trend. Default 5.
#>

param(
    [string]$ResultsDir = "",
    [int]$ShowTrend = 5
)

$ErrorActionPreference = "Stop"
$script:RepoRoot = Split-Path -Parent $PSScriptRoot
if (-not $script:RepoRoot -or -not (Test-Path $script:RepoRoot)) {
    $script:RepoRoot = Get-Location
}

function Write-Status($msg, $color = "White") { Write-Host $msg -ForegroundColor $color }

$csvPath = Join-Path $script:RepoRoot "tests\coverage-history.csv"

# ============ AUTO-DETECT RESULTS ============

if (-not $ResultsDir) {
    $latestDir = Join-Path $script:RepoRoot "tests\Run_Latest"
    if (Test-Path $latestDir) {
        $ResultsDir = $latestDir
    } else {
        $runsDir = Join-Path $script:RepoRoot "tests"
        $latestRun = Get-ChildItem $runsDir -Directory -Filter "Run_*" -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending | Select-Object -First 1
        if ($latestRun) {
            $ResultsDir = $latestRun.FullName
        }
    }
}

# ============ PARSE COMPARISON REPORT ============

function Parse-ComparisonReport([string]$reportPath) {
    if (-not (Test-Path $reportPath)) { return $null }

    $content = Get-Content $reportPath -Raw
    $entry = @{
        Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalFields = 0
        Match = 0
        Tolerance = 0
        NotImpl = 0
        Missing = 0
        Fail = 0
        Coverage = 0
    }

    # Parse "Summary: N match, N tolerance, N not-impl, N missing, N fail"
    if ($content -match 'Summary:\s*(\d+)\s*match,\s*(\d+)\s*tolerance,\s*(\d+)\s*not-impl,\s*(\d+)\s*missing,\s*(\d+)\s*fail') {
        $entry.Match = [int]$Matches[1]
        $entry.Tolerance = [int]$Matches[2]
        $entry.NotImpl = [int]$Matches[3]
        $entry.Missing = [int]$Matches[4]
        $entry.Fail = [int]$Matches[5]
    }

    # Parse "Coverage: XX.X% (N / M)"
    if ($content -match 'Coverage:\s*([\d.]+)%\s*\((\d+)\s*/\s*(\d+)\)') {
        $entry.Coverage = [double]$Matches[1]
        $entry.TotalFields = [int]$Matches[3]
    }

    # Extract date from report if available
    if ($content -match 'Generated:\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})') {
        $entry.Date = $Matches[1]
    }

    return $entry
}

# ============ RECORD ENTRY ============

Write-Status ""
Write-Status "============================================" "Cyan"
Write-Status " COVERAGE TRACKING" "Cyan"
Write-Status "============================================" "Cyan"
Write-Status ""

$reportPath = $null
if ($ResultsDir -and (Test-Path $ResultsDir)) {
    $reportPath = Join-Path $ResultsDir "comparison_report.txt"
}

if ($reportPath -and (Test-Path $reportPath)) {
    $entry = Parse-ComparisonReport $reportPath

    if ($entry -and $entry.TotalFields -gt 0) {
        Write-Status "Parsed comparison report from: $reportPath"
        Write-Status "  Fields: $($entry.TotalFields), Match: $($entry.Match), Tol: $($entry.Tolerance), NotImpl: $($entry.NotImpl), Missing: $($entry.Missing), Fail: $($entry.Fail)"
        Write-Status ""

        # Ensure CSV exists with header
        if (-not (Test-Path $csvPath)) {
            "date,total_fields,match,tolerance,not_implemented,missing,fail,coverage_pct" | Out-File $csvPath -Encoding UTF8
            Write-Status "Created coverage-history.csv" "DarkGray"
        }

        # Append entry
        $csvLine = "$($entry.Date),$($entry.TotalFields),$($entry.Match),$($entry.Tolerance),$($entry.NotImpl),$($entry.Missing),$($entry.Fail),$($entry.Coverage)"
        $csvLine | Out-File $csvPath -Append -Encoding UTF8
        Write-Status "Appended entry to $csvPath" "Green"
    } else {
        Write-Status "WARNING: Could not parse comparison report (no data found)" "Yellow"
        Write-Status "Run compare-qa-results.ps1 first to generate the report" "Yellow"
    }
} else {
    Write-Status "No comparison_report.txt found in results directory" "Yellow"
    Write-Status "Run compare-qa-results.ps1 first" "Yellow"
}

# ============ DISPLAY TREND ============

Write-Status ""
if (Test-Path $csvPath) {
    $csvData = Import-Csv $csvPath
    $recent = @($csvData | Select-Object -Last $ShowTrend)

    if ($recent.Count -eq 0) {
        Write-Status "No entries in coverage-history.csv" "Yellow"
        exit 0
    }

    Write-Status "--- Last $($recent.Count) runs ---" "Cyan"
    Write-Status ""

    # Header
    $header = "{0,-20} {1,6} {2,6} {3,4} {4,8} {5,7} {6,5} {7,8}" -f "Date", "Fields", "Match", "Tol", "NotImpl", "Missing", "Fail", "Coverage"
    Write-Status $header "DarkGray"
    Write-Status ("-" * 76) "DarkGray"

    foreach ($row in $recent) {
        $matchCount = [int]$row.match
        $tolCount = [int]$row.tolerance
        $notImplCount = [int]$row.not_implemented
        $missingCount = [int]$row.missing
        $failCount = [int]$row.fail
        $coveragePct = [double]$row.coverage_pct
        $totalFields = [int]$row.total_fields

        $line = "{0,-20} {1,6} {2,6} {3,4} {4,8} {5,7} {6,5} {7,7:N1}%" -f $row.date, $totalFields, $matchCount, $tolCount, $notImplCount, $missingCount, $failCount, $coveragePct

        $color = if ($failCount -gt 0) { "Red" } elseif ($coveragePct -ge 80) { "Green" } elseif ($coveragePct -ge 50) { "Yellow" } else { "White" }
        Write-Status $line $color
    }

    # Show delta if multiple runs
    if ($recent.Count -ge 2) {
        $first = $recent[0]
        $last = $recent[$recent.Count - 1]

        $matchDelta = [int]$last.match - [int]$first.match
        $notImplDelta = [int]$last.not_implemented - [int]$first.not_implemented
        $failDelta = [int]$last.fail - [int]$first.fail
        $coverageDelta = [double]$last.coverage_pct - [double]$first.coverage_pct

        Write-Status ""
        Write-Status "--- Trend (first -> last) ---" "Cyan"

        $matchSign = if ($matchDelta -ge 0) { "+" } else { "" }
        $matchColor = if ($matchDelta -gt 0) { "Green" } elseif ($matchDelta -lt 0) { "Red" } else { "White" }
        Write-Status "  Match:     ${matchSign}${matchDelta}" $matchColor

        $notImplSign = if ($notImplDelta -ge 0) { "+" } else { "" }
        $notImplColor = if ($notImplDelta -lt 0) { "Green" } elseif ($notImplDelta -gt 0) { "Red" } else { "White" }
        Write-Status "  NotImpl:   ${notImplSign}${notImplDelta}" $notImplColor

        $failSign = if ($failDelta -ge 0) { "+" } else { "" }
        $failColor = if ($failDelta -lt 0) { "Green" } elseif ($failDelta -gt 0) { "Red" } else { "White" }
        Write-Status "  Fail:      ${failSign}${failDelta}" $failColor

        $covSign = if ($coverageDelta -ge 0) { "+" } else { "" }
        $covColor = if ($coverageDelta -gt 0) { "Green" } elseif ($coverageDelta -lt 0) { "Red" } else { "White" }
        Write-Status ("  Coverage:  {0}{1:N1}%" -f $covSign, $coverageDelta) $covColor
    }
} else {
    Write-Status "No coverage-history.csv found yet" "Yellow"
    Write-Status "Run compare-qa-results.ps1 first, then re-run this script" "Yellow"
}

Write-Status ""
