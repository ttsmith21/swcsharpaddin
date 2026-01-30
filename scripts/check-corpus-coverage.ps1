<#
.SYNOPSIS
    Checks which Gold Standard test files exist and reports coverage.

.DESCRIPTION
    Compares files in GoldStandard_Inputs against the manifest to show:
    - Which files are present
    - Which files are missing
    - Coverage percentage by class

.PARAMETER InputDir
    Path to the inputs directory. Default: tests\GoldStandard_Inputs

.PARAMETER ManifestPath
    Path to the baseline manifest. Default: tests\GoldStandard_Baseline\manifest.json

.EXAMPLE
    .\check-corpus-coverage.ps1
#>

param (
    [string]$InputDir = "$PSScriptRoot\..\tests\GoldStandard_Inputs",
    [string]$ManifestPath = "$PSScriptRoot\..\tests\GoldStandard_Baseline\manifest.json"
)

$ErrorActionPreference = "Stop"

# Load manifest
if (!(Test-Path $ManifestPath)) {
    Write-Host "Manifest not found: $ManifestPath" -ForegroundColor Red
    exit 1
}

$manifest = Get-Content $ManifestPath -Raw | ConvertFrom-Json

# Get all expected files from manifest (skip _comment entries)
$expectedFiles = $manifest.files.PSObject.Properties.Name | Where-Object { $_ -notlike "_comment*" }

# Get existing files
$existingFiles = @()
if (Test-Path $InputDir) {
    $existingFiles = Get-ChildItem -Path $InputDir -File -Recurse |
        Where-Object { $_.Extension -in @('.sldprt', '.sldasm', '.step', '.igs', '.x_t', '.sat', '.dxf') } |
        Select-Object -ExpandProperty Name
}

# Classify by prefix
$classes = @{
    "A" = @{ Name = "Invalid/Problem"; Priority = "High"; Expected = @(); Present = @(); Missing = @() }
    "B" = @{ Name = "Sheet Metal"; Priority = "High"; Expected = @(); Present = @(); Missing = @() }
    "C" = @{ Name = "Tube & Structural"; Priority = "High"; Expected = @(); Present = @(); Missing = @() }
    "D" = @{ Name = "Non-Convertible"; Priority = "Medium"; Expected = @(); Present = @(); Missing = @() }
    "E" = @{ Name = "Edge Cases"; Priority = "Medium"; Expected = @(); Present = @(); Missing = @() }
    "F" = @{ Name = "Assemblies"; Priority = "Medium"; Expected = @(); Present = @(); Missing = @() }
    "G" = @{ Name = "Configurations"; Priority = "Low"; Expected = @(); Present = @(); Missing = @() }
    "H" = @{ Name = "File Formats"; Priority = "Low"; Expected = @(); Present = @(); Missing = @() }
}

# Categorize expected files
foreach ($file in $expectedFiles) {
    $prefix = $file.Substring(0, 1).ToUpper()
    if ($classes.ContainsKey($prefix)) {
        $classes[$prefix].Expected += $file
    }
}

# Check presence
foreach ($classKey in $classes.Keys) {
    foreach ($file in $classes[$classKey].Expected) {
        if ($existingFiles -contains $file) {
            $classes[$classKey].Present += $file
        } else {
            $classes[$classKey].Missing += $file
        }
    }
}

# Report
Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " GOLD STANDARD CORPUS COVERAGE REPORT" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$totalExpected = 0
$totalPresent = 0

foreach ($classKey in ($classes.Keys | Sort-Object)) {
    $c = $classes[$classKey]
    $expected = $c.Expected.Count
    $present = $c.Present.Count
    $pct = if ($expected -gt 0) { [math]::Round(($present / $expected) * 100) } else { 0 }

    $totalExpected += $expected
    $totalPresent += $present

    $color = if ($pct -eq 100) { "Green" } elseif ($pct -gt 0) { "Yellow" } else { "Gray" }
    $priorityColor = switch ($c.Priority) { "High" { "Red" } "Medium" { "Yellow" } "Low" { "Gray" } }

    Write-Host -NoNewline "  Class $classKey - $($c.Name) "
    Write-Host -NoNewline "[$($c.Priority)]" -ForegroundColor $priorityColor
    Write-Host ""

    $bar = "[" + ("█" * [math]::Floor($pct / 10)) + ("░" * (10 - [math]::Floor($pct / 10))) + "]"
    Write-Host "    $bar $present/$expected ($pct%)" -ForegroundColor $color

    if ($c.Missing.Count -gt 0 -and $c.Missing.Count -le 3) {
        Write-Host "    Missing: $($c.Missing -join ', ')" -ForegroundColor Gray
    } elseif ($c.Missing.Count -gt 3) {
        Write-Host "    Missing: $($c.Missing[0..2] -join ', '), +$($c.Missing.Count - 3) more" -ForegroundColor Gray
    }
    Write-Host ""
}

$totalPct = if ($totalExpected -gt 0) { [math]::Round(($totalPresent / $totalExpected) * 100) } else { 0 }

Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  TOTAL: $totalPresent / $totalExpected files ($totalPct%)" -ForegroundColor $(if ($totalPct -eq 100) { "Green" } else { "Yellow" })
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Runnable subset check
if ($totalPresent -gt 0) {
    Write-Host "Ready to run tests on $totalPresent available files." -ForegroundColor Green
    Write-Host "Run: .\scripts\run-gold-standard-tests.ps1" -ForegroundColor Gray
} else {
    Write-Host "No test files found. Add files to: tests\GoldStandard_Inputs\" -ForegroundColor Yellow
}

# Export missing files list
$missingFile = "$PSScriptRoot\..\tests\missing-files.txt"
$allMissing = @()
foreach ($classKey in ($classes.Keys | Sort-Object)) {
    $allMissing += $classes[$classKey].Missing
}
if ($allMissing.Count -gt 0) {
    $allMissing | Set-Content $missingFile
    Write-Host ""
    Write-Host "Missing files list saved to: tests\missing-files.txt" -ForegroundColor Gray
}
