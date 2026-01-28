<#
.SYNOPSIS
    Quick incremental build (no clean) for fast iteration.
.DESCRIPTION
    Runs MSBuild without cleaning, outputs errors to console and build-errors.txt.
    Use this for rapid fix-compile cycles.
#>

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not $projectRoot) { $projectRoot = (Get-Location).Path }

# Find MSBuild
$msbuild = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" `
    -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\amd64\MSBuild.exe | Select-Object -First 1

if (-not $msbuild -or -not (Test-Path $msbuild)) {
    Write-Error "MSBuild not found"
    exit 1
}

$csproj = Join-Path $projectRoot "swcsharpaddin.csproj"
$errorLog = Join-Path $projectRoot "build-errors.txt"

Write-Host "=== Quick Build ===" -ForegroundColor Cyan
Write-Host "Project: $csproj"

# Run incremental build, capture output
$output = & $msbuild $csproj /p:Configuration=Debug /v:minimal /nologo 2>&1

# Extract errors and warnings
$errors = $output | Select-String "error CS\d+"
$warnings = $output | Select-String "warning CS\d+" | Select-Object -First 10

# Write full errors to file
$allErrors = $output | Select-String "error CS"
if ($allErrors) {
    $allErrors | Out-File -FilePath $errorLog -Encoding UTF8
    Write-Host "`nErrors written to: build-errors.txt" -ForegroundColor Yellow
}

# Count unique errors
$uniqueErrors = $errors | ForEach-Object {
    if ($_ -match "(error CS\d+)") { $matches[1] }
} | Sort-Object -Unique

$errorCount = ($errors | Measure-Object).Count
$uniqueCount = ($uniqueErrors | Measure-Object).Count

if ($errorCount -eq 0) {
    Write-Host "`nBUILD: SUCCESS" -ForegroundColor Green
    if (Test-Path $errorLog) { Remove-Item $errorLog -Force }
    exit 0
} else {
    Write-Host "`nBUILD: FAILED" -ForegroundColor Red
    Write-Host "  Total errors: $errorCount ($uniqueCount unique)" -ForegroundColor Red
    Write-Host "`nUnique error codes:" -ForegroundColor Yellow
    $uniqueErrors | ForEach-Object { Write-Host "  $_" }
    Write-Host "`nFirst 10 errors:" -ForegroundColor Yellow
    $errors | Select-Object -First 10 | ForEach-Object { Write-Host "  $_" }
    exit 1
}
