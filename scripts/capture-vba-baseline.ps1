<#
.SYNOPSIS
    Captures custom properties from files processed by VBA macro as a baseline.

.DESCRIPTION
    After running the VBA macro on your test corpus:
    1. This script writes a config for the add-in to read properties only
    2. The add-in extracts current custom properties (written by VBA)
    3. Output becomes the "VBA baseline" to compare C# results against

    Workflow:
    1. Copy test files to a working directory
    2. Run VBA macro on all files (manually in SolidWorks)
    3. Run this script
    4. Click "Run QA" in SolidWorks (reads properties, doesn't process)
    5. Script generates vba-baseline.json

.PARAMETER InputDir
    Path to files processed by VBA. Default: tests\VBA_Processed

.EXAMPLE
    .\capture-vba-baseline.ps1 -InputDir "C:\TestFiles\VBA_Output"
#>

param (
    [string]$InputDir = "$PSScriptRoot\..\tests\VBA_Processed"
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " VBA BASELINE CAPTURE" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

if (!(Test-Path $InputDir)) {
    Write-Host "Input directory not found: $InputDir" -ForegroundColor Red
    Write-Host ""
    Write-Host "Expected workflow:" -ForegroundColor Yellow
    Write-Host "  1. Create directory: tests\VBA_Processed"
    Write-Host "  2. Copy test files there"
    Write-Host "  3. Run VBA macro on all files in SolidWorks"
    Write-Host "  4. Run this script again"
    exit 1
}

$files = Get-ChildItem -Path $InputDir -Filter "*.sldprt" -ErrorAction SilentlyContinue
if ($files.Count -eq 0) {
    Write-Host "No .sldprt files found in: $InputDir" -ForegroundColor Red
    exit 1
}

Write-Host "Found $($files.Count) files to capture baseline from"
Write-Host ""

# Create config for property-read-only mode
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputDir = "$PSScriptRoot\..\tests\VBA_Baseline_$timestamp"
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$configPath = "C:\Temp\nm_qa_config.json"
$resultFile = "$outputDir\vba-baseline-raw.json"

$config = @{
    InputPath = $InputDir
    OutputPath = $resultFile
    Mode = "ReadPropertiesOnly"  # Signal to add-in (if we implement this mode)
}

if (!(Test-Path "C:\Temp")) {
    New-Item -ItemType Directory -Force -Path "C:\Temp" | Out-Null
}

$config | ConvertTo-Json | Set-Content $configPath -Encoding UTF8

Write-Host "Config written to: $configPath" -ForegroundColor Gray
Write-Host "Results will be saved to: $resultFile" -ForegroundColor Gray
Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Yellow
Write-Host " NEXT STEPS:" -ForegroundColor Yellow
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Yellow
Write-Host ""
Write-Host "  1. Open SolidWorks"
Write-Host "  2. Click 'Run QA' from the add-in toolbar"
Write-Host "  3. Wait for completion"
Write-Host "  4. Run: .\scripts\convert-vba-to-manifest.ps1 -ResultsFile '$resultFile'"
Write-Host ""
Write-Host "This will generate a manifest.json with VBA's values as the expected baseline."
Write-Host ""

# Wait for results
Write-Host "Waiting for results (5 minute timeout)..." -ForegroundColor Gray
$timeout = (Get-Date).AddMinutes(5)

while (!(Test-Path $resultFile)) {
    if ((Get-Date) -gt $timeout) {
        Write-Host ""
        Write-Host "Timeout waiting for results. Run 'Run QA' in SolidWorks." -ForegroundColor Red
        exit 1
    }
    Start-Sleep -Seconds 2
    Write-Host "." -NoNewline
}

Write-Host ""
Write-Host ""
Write-Host "Results captured!" -ForegroundColor Green
Write-Host "Raw data: $resultFile"
Write-Host ""
Write-Host "To generate manifest, run:" -ForegroundColor Cyan
Write-Host "  .\scripts\create-baseline.ps1 -ResultsFile '$resultFile'"
