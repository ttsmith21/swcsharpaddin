<#
.SYNOPSIS
    Runs SolidWorks add-in commands in headless mode using NM.BatchRunner.

.DESCRIPTION
    Launches NM.BatchRunner.exe which starts SolidWorks, runs QA or Pipeline,
    and closes SolidWorks when done.

.PARAMETER Mode
    The mode to run: "qa" or "pipeline"

.PARAMETER FilePath
    For pipeline mode: path to the SLDPRT file to process

.EXAMPLE
    .\run-headless.ps1 -Mode qa

.EXAMPLE
    .\run-headless.ps1 -Mode pipeline -FilePath "C:\Parts\MyPart.SLDPRT"
#>

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("qa", "pipeline")]
    [string]$Mode,

    [string]$FilePath
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDir
$batchRunnerPath = Join-Path $projectRoot "src\NM.BatchRunner\bin\Debug\NM.BatchRunner.exe"

# Verify BatchRunner exists
if (!(Test-Path $batchRunnerPath)) {
    Write-Error "NM.BatchRunner.exe not found. Run build first: .\scripts\build-and-test.ps1"
    exit 1
}

# Build arguments
if ($Mode -eq "qa") {
    $arguments = "--qa"
} else {
    if ([string]::IsNullOrEmpty($FilePath)) {
        Write-Error "Pipeline mode requires -FilePath parameter"
        exit 1
    }
    if (!(Test-Path $FilePath)) {
        Write-Error "File not found: $FilePath"
        exit 1
    }
    $arguments = "--pipeline --file `"$FilePath`""
}

Write-Host "Running: NM.BatchRunner.exe $arguments" -ForegroundColor Cyan
Write-Host ""

# Run BatchRunner
$startTime = Get-Date
& $batchRunnerPath $arguments.Split(' ')
$exitCode = $LASTEXITCODE
$duration = (Get-Date) - $startTime

Write-Host ""
Write-Host "Completed in $($duration.TotalSeconds.ToString('F1')) seconds" -ForegroundColor Green
Write-Host "Exit code: $exitCode" -ForegroundColor $(if ($exitCode -eq 0) { "Green" } else { "Yellow" })

exit $exitCode
