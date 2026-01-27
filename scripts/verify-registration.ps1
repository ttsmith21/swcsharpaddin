<#
.SYNOPSIS
    Verifies that the SolidWorks add-in is registered correctly.

.DESCRIPTION
    Checks the Windows registry for the add-in registration under:
    - HKLM:\SOFTWARE\SolidWorks\Addins\{GUID}
    - HKCU:\SOFTWARE\SolidWorks\Addins\{GUID}

    Also checks COM registration under HKCR\CLSID.

.PARAMETER Fix
    If specified and registration is missing, attempts to register using regasm.
    Requires Administrator privileges.
#>
param(
    [switch]$Fix
)

$ErrorActionPreference = "Continue"
$projectRoot = Split-Path -Parent $PSScriptRoot
$addInGuid = "{D5355548-9569-4381-9939-5D14252A3E47}"
$dllPath = Join-Path $projectRoot "bin\Debug\swcsharpaddin.dll"

Write-Host "=== verify-registration ===" -ForegroundColor Cyan
Write-Host "Add-in GUID: $addInGuid"
Write-Host "DLL Path: $dllPath"
Write-Host ""

$allGood = $true

# Check if DLL exists
if (-not (Test-Path $dllPath)) {
    Write-Host "DLL: NOT FOUND" -ForegroundColor Red
    Write-Host "  Run build-and-test.ps1 first"
    $allGood = $false
} else {
    Write-Host "DLL: EXISTS" -ForegroundColor Green
    $dllInfo = Get-Item $dllPath
    Write-Host "  Size: $([math]::Round($dllInfo.Length / 1KB, 1)) KB"
    Write-Host "  Modified: $($dllInfo.LastWriteTime)"
}
Write-Host ""

# Check SolidWorks add-in registration (HKLM)
$swRegPathLM = "HKLM:\SOFTWARE\SolidWorks\Addins\$addInGuid"
if (Test-Path $swRegPathLM) {
    Write-Host "SW REGISTRATION (HKLM): FOUND" -ForegroundColor Green
    $props = Get-ItemProperty $swRegPathLM -ErrorAction SilentlyContinue
    if ($props) {
        Write-Host "  Description: $($props.Description)"
        Write-Host "  Title: $($props.Title)"
    }
} else {
    Write-Host "SW REGISTRATION (HKLM): NOT FOUND" -ForegroundColor Yellow
    $allGood = $false
}

# Check SolidWorks add-in registration (HKCU)
$swRegPathCU = "HKCU:\SOFTWARE\SolidWorks\Addins\$addInGuid"
if (Test-Path $swRegPathCU) {
    Write-Host "SW REGISTRATION (HKCU): FOUND" -ForegroundColor Green
} else {
    Write-Host "SW REGISTRATION (HKCU): NOT FOUND" -ForegroundColor Yellow
}
Write-Host ""

# Check COM registration
$comRegPath = "Registry::HKEY_CLASSES_ROOT\CLSID\$addInGuid"
if (Test-Path $comRegPath) {
    Write-Host "COM REGISTRATION: FOUND" -ForegroundColor Green
    $inprocServer = Get-ItemProperty "$comRegPath\InprocServer32" -ErrorAction SilentlyContinue
    if ($inprocServer) {
        Write-Host "  CodeBase: $($inprocServer.CodeBase)"
        Write-Host "  Assembly: $($inprocServer.Assembly)"
    }
} else {
    Write-Host "COM REGISTRATION: NOT FOUND" -ForegroundColor Yellow
    $allGood = $false
}
Write-Host ""

# Summary
Write-Host "=== Summary ===" -ForegroundColor Cyan
if ($allGood) {
    Write-Host "ADD-IN REGISTERED SUCCESSFULLY" -ForegroundColor Green
    exit 0
} else {
    Write-Host "ADD-IN REGISTRATION INCOMPLETE" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "To register manually (requires Administrator):" -ForegroundColor Gray
    Write-Host "  C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe /codebase `"$dllPath`"" -ForegroundColor Gray

    if ($Fix) {
        Write-Host ""
        Write-Host "Attempting registration..." -ForegroundColor Yellow
        $regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
        if (Test-Path $regasm) {
            $result = & $regasm /codebase $dllPath 2>&1 | Out-String
            Write-Host $result
        } else {
            Write-Host "ERROR: RegAsm not found at expected path" -ForegroundColor Red
        }
    }

    exit 1
}
