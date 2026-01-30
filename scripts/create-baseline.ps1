<#
.SYNOPSIS
    Creates or updates the Gold Standard baseline from a known-good QA run.

.DESCRIPTION
    Takes the results.json from a QA run and generates/updates the baseline
    manifest.json with expected values.

.PARAMETER ResultsFile
    Path to a results.json from a previous QA run.

.PARAMETER BaselineDir
    Path to the baseline directory. Default: tests\GoldStandard_Baseline

.PARAMETER Merge
    If set, merge with existing baseline instead of replacing.

.EXAMPLE
    .\create-baseline.ps1 -ResultsFile ..\tests\Run_20250128_143000\results.json

.EXAMPLE
    .\create-baseline.ps1 -ResultsFile ..\tests\Run_20250128_143000\results.json -Merge
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$ResultsFile,

    [string]$BaselineDir = "$PSScriptRoot\..\tests\GoldStandard_Baseline",

    [switch]$Merge
)

$ErrorActionPreference = "Stop"

function Write-Status {
    param([string]$Message, [string]$Color = "Cyan")
    Write-Host "[Baseline] $Message" -ForegroundColor $Color
}

# ============================================================================
# 1. Load Results
# ============================================================================
if (!(Test-Path $ResultsFile)) {
    Write-Host "[Baseline] ERROR: Results file not found: $ResultsFile" -ForegroundColor Red
    exit 1
}

$results = Get-Content $ResultsFile -Raw | ConvertFrom-Json
Write-Status "Loaded $($results.Results.Count) results from $ResultsFile"

# ============================================================================
# 2. Load Existing Baseline (if merging)
# ============================================================================
$baselineFile = "$BaselineDir\manifest.json"
$baseline = $null

if ($Merge -and (Test-Path $baselineFile)) {
    $baseline = Get-Content $baselineFile -Raw | ConvertFrom-Json
    Write-Status "Merging with existing baseline"
} else {
    $baseline = [PSCustomObject]@{
        '$schema' = "./manifest-schema.json"
        version = "1.0"
        description = "Gold Standard expected outcomes for regression testing"
        generatedFrom = $ResultsFile
        generatedAt = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ")
        defaultTolerances = [PSCustomObject]@{
            thickness = 0.001
            mass = 0.01
            dimensions = 0.001
            cost = 0.05
        }
        files = [PSCustomObject]@{}
    }
    Write-Status "Creating new baseline"
}

# ============================================================================
# 3. Process Results into Baseline Entries
# ============================================================================
Write-Status "Processing results..."

$addedCount = 0
$updatedCount = 0

foreach ($res in $results.Results) {
    $fileName = $res.FileName

    # Skip if status is Error (something went wrong, not a valid baseline)
    if ($res.Status -eq "Error") {
        Write-Host "   [SKIP] $fileName - Error status" -ForegroundColor Yellow
        continue
    }

    $entry = [ordered]@{}

    # Basic info
    $entry["shouldPass"] = ($res.Status -eq "Success")

    if ($res.Status -ne "Success" -and $res.Message) {
        # Extract failure reason category
        $reason = $res.Message
        if ($reason -like "*Multi-body*") { $reason = "Multi-body" }
        elseif ($reason -like "*No solid body*") { $reason = "No solid body" }
        elseif ($reason -like "*material*") { $reason = "Missing material" }
        $entry["expectedFailureReason"] = $reason
    }

    # Classification
    if ($res.Classification) {
        $entry["expectedClassification"] = $res.Classification
    }

    # Geometry (only include if present)
    if ($null -ne $res.Thickness_in -and $res.Thickness_in -gt 0) {
        $entry["expectedThickness_in"] = [Math]::Round($res.Thickness_in, 6)
    }

    if ($null -ne $res.Mass_lb -and $res.Mass_lb -gt 0) {
        $entry["expectedMass_lb"] = [Math]::Round($res.Mass_lb, 4)
    }

    # Sheet metal
    if ($null -ne $res.BendCount) {
        $entry["expectedBendCount"] = $res.BendCount
    }

    # Tube
    if ($null -ne $res.TubeOD_in -and $res.TubeOD_in -gt 0) {
        $entry["expectedTubeOD_in"] = [Math]::Round($res.TubeOD_in, 4)
    }

    if ($null -ne $res.TubeWall_in -and $res.TubeWall_in -gt 0) {
        $entry["expectedTubeWall_in"] = [Math]::Round($res.TubeWall_in, 4)
    }

    if ($null -ne $res.TubeLength_in -and $res.TubeLength_in -gt 0) {
        $entry["expectedTubeLength_in"] = [Math]::Round($res.TubeLength_in, 4)
    }

    # Add or update
    $exists = $baseline.files.PSObject.Properties.Name -contains $fileName
    if ($exists) {
        $baseline.files.$fileName = [PSCustomObject]$entry
        $updatedCount++
        Write-Host "   [UPDATE] $fileName" -ForegroundColor Yellow
    } else {
        $baseline.files | Add-Member -MemberType NoteProperty -Name $fileName -Value ([PSCustomObject]$entry)
        $addedCount++
        Write-Host "   [ADD] $fileName" -ForegroundColor Green
    }
}

# ============================================================================
# 4. Write Baseline
# ============================================================================
# Ensure directory exists
if (!(Test-Path $BaselineDir)) {
    New-Item -ItemType Directory -Force -Path $BaselineDir | Out-Null
}

# Update metadata
$baseline.generatedFrom = $ResultsFile
$baseline.generatedAt = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ")

# Convert to JSON and write
$json = $baseline | ConvertTo-Json -Depth 10
Set-Content -Path $baselineFile -Value $json -Encoding UTF8

Write-Status "Baseline written to: $baselineFile" -Color Green
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  Added:   $addedCount"
Write-Host "  Updated: $updatedCount"
Write-Host "  Total:   $($baseline.files.PSObject.Properties.Count)"
Write-Host ""
Write-Status "Review the manifest.json and adjust tolerances as needed" -Color Yellow
