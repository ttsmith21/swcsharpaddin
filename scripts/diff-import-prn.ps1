<#
.SYNOPSIS
    Compares VBA baseline Import.prn with C# generated Import_CSharp.prn.

.DESCRIPTION
    Parses both Import.prn files section-by-section (IM, ML, PS, RT, RN),
    groups records by part key, and reports per-part field-level differences
    with numeric tolerance support.

.PARAMETER VbaPath
    Path to VBA baseline Import.prn. Auto-detects from GoldStandard_Baseline.

.PARAMETER CSharpPath
    Path to C# generated Import_CSharp.prn. Auto-detects from latest Run_*.

.PARAMETER Tolerance
    Numeric tolerance for comparisons. Default 0.05 (5%)
#>

param(
    [string]$VbaPath = "",
    [string]$CSharpPath = "",
    [double]$Tolerance = 0.05
)

$ErrorActionPreference = "Stop"
$script:RepoRoot = Split-Path -Parent $PSScriptRoot
if (-not $script:RepoRoot -or -not (Test-Path $script:RepoRoot)) {
    $script:RepoRoot = Get-Location
}

function Write-Status($msg, $color = "White") { Write-Host $msg -ForegroundColor $color }

# ============ TOKENIZER ============
function Tokenize-PrnLine([string]$line) {
    $tokens = @()
    $current = ""
    $inQuotes = $false

    for ($i = 0; $i -lt $line.Length; $i++) {
        $ch = $line[$i]

        if ($ch -eq '"') {
            $inQuotes = -not $inQuotes
            continue
        }

        if (-not $inQuotes -and ($ch -eq ' ' -or $ch -eq "`t")) {
            if ($current.Length -gt 0) {
                $tokens += $current
                $current = ""
            }
            continue
        }

        $current += $ch
    }

    if ($current.Length -gt 0) {
        $tokens += $current
    }

    return ,$tokens
}

# ============ PARSER ============
function Parse-ImportPrn([string]$FilePath) {
    $lines = @(Get-Content $FilePath)
    $result = @{
        IM = @()    # Item Master
        ML = @()    # Material Locations
        PS = @()    # Product Structure / BOM + Material relationships
        RT = @()    # Routing
        RN = @()    # Routing Notes
    }

    $currentSection = $null
    $fieldNames = @()
    $inDataSection = $false

    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            $inDataSection = $false
            $currentSection = $null
            continue
        }

        # Section header: DECL(XX) [ADD] field1 field2 ...
        if ($line -match '^\s*DECL\((\w+)\)') {
            $sectionCode = $Matches[1]
            $inDataSection = $false

            # Extract field names (skip DECL(XX) and optional ADD)
            $headerTokens = @(Tokenize-PrnLine $line)
            $startIdx = 1
            if ($headerTokens.Count -gt 1 -and $headerTokens[1] -eq "ADD") {
                $startIdx = 2
            }
            $fieldNames = @()
            for ($i = $startIdx; $i -lt $headerTokens.Count; $i++) {
                $fieldNames += $headerTokens[$i]
            }

            $currentSection = $sectionCode
            continue
        }

        if ($line.Trim() -eq "END") {
            $inDataSection = $true
            continue
        }

        if (-not $inDataSection -or $null -eq $currentSection) { continue }

        # Parse data row
        $tokens = @(Tokenize-PrnLine $line)
        if ($tokens.Count -eq 0) { continue }

        $record = @{}
        $maxFields = [Math]::Min($fieldNames.Count, $tokens.Count)
        for ($i = 0; $i -lt $maxFields; $i++) {
            $record[$fieldNames[$i]] = $tokens[$i]
        }

        # Store in appropriate section
        if ($result.ContainsKey($currentSection)) {
            $result[$currentSection] += $record
        } else {
            $result[$currentSection] = @($record)
        }
    }

    return $result
}

# ============ COMPARISON ============
function Compare-Numeric($a, $b, $tol) {
    if ($null -eq $a -or $null -eq $b) { return $false }
    try {
        $va = [double]$a; $vb = [double]$b
    } catch { return $false }
    if ($vb -eq 0) { return [Math]::Abs($va) -lt 0.001 }
    $diff = [Math]::Abs($va - $vb)
    $threshold = [Math]::Max([Math]::Abs($vb * $tol), 0.001)
    return $diff -le $threshold
}

function Is-NumericField([string]$fieldName) {
    $numericFields = @(
        "IM-STD-LOT", "IM-STD-MAT", "IM-TYPE", "IM-CLASS",
        "IM-SAVE-DEMAND-SW", "IM-PLANNER", "IM-STOCK-SW", "IM-PLAN-SW", "IM-ISSUE-SW",
        "PS-QTY-P", "PS-PIECE-NO", "PS-DIM-1", "PS-ISSUE-SW",
        "PS-BFLOCATION-SW", "PS-BFQTY-SW", "PS-BFZEROQTY-SW", "PS-OP-NUM",
        "RT-OP-NUM", "RT-SETUP", "RT-RUN-STD", "RT-MULT-SEQ",
        "RN-OP-NUM", "RN-LINE-NO"
    )
    return $numericFields -contains $fieldName
}

function Get-PartKey($record) {
    foreach ($keyField in @("IM-KEY", "ML-IMKEY", "PS-PARENT-KEY", "PS-SUBORD-KEY", "RT-ITEM-KEY", "RN-ITEM-KEY")) {
        if ($record.ContainsKey($keyField) -and $record[$keyField]) {
            return $record[$keyField]
        }
    }
    return $null
}

function Compare-Records($expRecord, $actRecord, $tolerance) {
    $diffs = @()
    $allFields = @($expRecord.Keys) + @($actRecord.Keys) | Select-Object -Unique

    foreach ($field in $allFields) {
        $expVal = if ($expRecord.ContainsKey($field)) { $expRecord[$field] } else { "(missing)" }
        $actVal = if ($actRecord.ContainsKey($field)) { $actRecord[$field] } else { "(missing)" }

        if ($expVal -eq $actVal) { continue }

        # Try numeric comparison with tolerance
        if ((Is-NumericField $field)) {
            if ((Compare-Numeric $actVal $expVal $tolerance)) {
                $diffs += @{ Field = $field; Status = "TOLERANCE"; Expected = $expVal; Actual = $actVal }
                continue
            }
        }

        $diffs += @{ Field = $field; Status = "DIFF"; Expected = $expVal; Actual = $actVal }
    }

    return ,$diffs
}

# ============ MAIN ============

Write-Status ""
Write-Status "============================================" "Cyan"
Write-Status " IMPORT.PRN STRUCTURAL DIFF" "Cyan"
Write-Status "============================================" "Cyan"
Write-Status ""

# Auto-detect VBA baseline
if (-not $VbaPath) {
    $baselineDir = Join-Path $script:RepoRoot "tests\GoldStandard_Baseline"
    $candidates = @(Get-ChildItem $baselineDir -Filter "Import_*.prn" -ErrorAction SilentlyContinue)
    if ($candidates.Count -gt 0) {
        $VbaPath = $candidates[0].FullName
    }
}

# Auto-detect C# output
if (-not $CSharpPath) {
    # Try Run_Latest first
    $latestDir = Join-Path $script:RepoRoot "tests\Run_Latest"
    $csharpFile = Join-Path $latestDir "Import_CSharp.prn"
    if (Test-Path $csharpFile) {
        $CSharpPath = $csharpFile
    } else {
        # Find most recent Run_* folder
        $runsDir = Join-Path $script:RepoRoot "tests"
        $latestRun = Get-ChildItem $runsDir -Directory -Filter "Run_*" -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending | Select-Object -First 1
        if ($latestRun) {
            $CSharpPath = Join-Path $latestRun.FullName "Import_CSharp.prn"
        }
    }
}

if (-not $VbaPath -or -not (Test-Path $VbaPath)) {
    Write-Status "ERROR: VBA baseline Import.prn not found: $VbaPath" "Red"
    Write-Status "Expected in tests\GoldStandard_Baseline\Import_*.prn" "Yellow"
    exit 1
}
if (-not $CSharpPath -or -not (Test-Path $CSharpPath)) {
    Write-Status "ERROR: C# Import_CSharp.prn not found: $CSharpPath" "Red"
    Write-Status "Run /qa first to generate Import_CSharp.prn" "Yellow"
    exit 1
}

Write-Status "VBA Baseline: $VbaPath"
Write-Status "C# Output:    $CSharpPath"
Write-Status ""

try {
    $vbaSections = Parse-ImportPrn $VbaPath
    $csharpSections = Parse-ImportPrn $CSharpPath
} catch {
    Write-Status "ERROR: Failed to parse .prn files: $_" "Red"
    exit 1
}

$totalMatch = 0; $totalTol = 0; $totalDiff = 0; $totalMissing = 0
$sectionStats = @{}

foreach ($section in @("IM", "ML", "PS", "RT", "RN")) {
    $vbaRecords = @()
    if ($vbaSections.ContainsKey($section)) { $vbaRecords = @($vbaSections[$section]) }
    $csRecords = @()
    if ($csharpSections.ContainsKey($section)) { $csRecords = @($csharpSections[$section]) }

    $sMatch = 0; $sTol = 0; $sDiff = 0; $sMissing = 0

    Write-Status "--- Section: $section ---" "Cyan"
    Write-Status "  VBA records: $($vbaRecords.Count), C# records: $($csRecords.Count)"

    if ($vbaRecords.Count -ne $csRecords.Count) {
        $countDiff = $csRecords.Count - $vbaRecords.Count
        $sign = if ($countDiff -gt 0) { "+" } else { "" }
        Write-Status "  Record count difference: ${sign}${countDiff}" "Yellow"
    }

    # Compare by index (records should be in same order)
    $maxRecords = [Math]::Max($vbaRecords.Count, $csRecords.Count)
    for ($i = 0; $i -lt $maxRecords; $i++) {
        if ($i -ge $vbaRecords.Count) {
            $partKey = Get-PartKey $csRecords[$i]
            Write-Status "  [EXTRA] Row $i ($partKey): only in C# output" "Yellow"
            $sMissing++
            continue
        }
        if ($i -ge $csRecords.Count) {
            $partKey = Get-PartKey $vbaRecords[$i]
            Write-Status "  [MISSING] Row $i ($partKey): only in VBA baseline" "Yellow"
            $sMissing++
            continue
        }

        $diffs = @(Compare-Records $vbaRecords[$i] $csRecords[$i] $Tolerance)
        $partKey = Get-PartKey $vbaRecords[$i]

        if ($diffs.Count -eq 0) {
            $sMatch++
        } else {
            $tolCount = @($diffs | Where-Object { $_.Status -eq "TOLERANCE" }).Count
            $diffCount = @($diffs | Where-Object { $_.Status -eq "DIFF" }).Count

            if ($diffCount -gt 0) {
                $sDiff++
                $color = "Red"
            } else {
                $sTol++
                $color = "DarkYellow"
            }

            Write-Status "  [$($partKey)] $diffCount diffs, $tolCount tolerance" $color
            foreach ($d in $diffs) {
                $dColor = if ($d.Status -eq "TOLERANCE") { "DarkYellow" } else { "Red" }
                Write-Status ("    [{0,-9}] {1}: vba={2}, csharp={3}" -f $d.Status, $d.Field, $d.Expected, $d.Actual) $dColor
            }
        }
    }

    $totalMatch += $sMatch; $totalTol += $sTol; $totalDiff += $sDiff; $totalMissing += $sMissing
    $sectionStats[$section] = @{ Match = $sMatch; Tolerance = $sTol; Diff = $sDiff; Missing = $sMissing }
    Write-Status ""
}

# Summary
Write-Status "============================================" "Cyan"
Write-Status " SUMMARY" "Cyan"
Write-Status "============================================" "Cyan"
Write-Status ""

foreach ($section in @("IM", "ML", "PS", "RT", "RN")) {
    if ($sectionStats.ContainsKey($section)) {
        $s = $sectionStats[$section]
        $sColor = if ($s.Diff -gt 0) { "Red" } elseif ($s.Missing -gt 0 -or $s.Tolerance -gt 0) { "Yellow" } else { "Green" }
        Write-Status ("  {0}: {1} match, {2} tolerance, {3} diff, {4} missing" -f $section, $s.Match, $s.Tolerance, $s.Diff, $s.Missing) $sColor
    }
}

$totalRecords = $totalMatch + $totalTol + $totalDiff + $totalMissing
Write-Status ""
Write-Status "  Total records: $totalRecords"
Write-Status "  Match:     $totalMatch" "Green"
Write-Status "  Tolerance: $totalTol" $(if ($totalTol -gt 0) { "DarkYellow" } else { "Green" })
Write-Status "  Diff:      $totalDiff" $(if ($totalDiff -gt 0) { "Red" } else { "Green" })
Write-Status "  Missing:   $totalMissing" $(if ($totalMissing -gt 0) { "Yellow" } else { "Green" })

$pct = if ($totalRecords -gt 0) { [Math]::Round(($totalMatch + $totalTol) / $totalRecords * 100, 1) } else { 0 }
Write-Status ""
Write-Status "  Match rate: $pct% ($($totalMatch + $totalTol) / $totalRecords)" $(if ($pct -ge 80) { "Green" } elseif ($pct -ge 50) { "Yellow" } else { "Red" })
Write-Status ""

# Write report
$reportDir = Split-Path $CSharpPath
$reportPath = Join-Path $reportDir "import_prn_diff.txt"
$report = @()
$report += "Import.prn Diff Report"
$report += "======================"
$report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$report += "VBA:    $VbaPath"
$report += "C#:     $CSharpPath"
$report += ""
foreach ($section in @("IM", "ML", "PS", "RT", "RN")) {
    if ($sectionStats.ContainsKey($section)) {
        $s = $sectionStats[$section]
        $report += "$section`: $($s.Match) match, $($s.Tolerance) tolerance, $($s.Diff) diff, $($s.Missing) missing"
    }
}
$report += ""
$report += "Match rate: $pct% ($($totalMatch + $totalTol) / $totalRecords)"
$report | Out-File $reportPath -Encoding UTF8
Write-Status "Report: $reportPath"

if ($totalDiff -gt 0) { exit 1 } else { exit 0 }
