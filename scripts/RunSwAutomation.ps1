param(
    [string]$Root = 'C:\SWTests',
    [string[]]$Extensions = @('.sldprt', '.stp', '.step'),
    [string]$AddInGuid = '{d5355548-9569-4381-9939-5d14252a3e47}',
    [string]$AddInPath,
    [string]$SwExePath = 'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe'
)

$ErrorActionPreference = 'Stop'

function Write-Info($msg) { Write-Host (Get-Date -Format 'HH:mm:ss') ": $msg" }

function Register-AddIn([string]$dllPath) {
    if (-not (Test-Path $dllPath)) { throw "Add-in DLL not found: $dllPath" }
    $regasm = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
    if (-not (Test-Path $regasm)) { throw "RegAsm.exe not found at $regasm" }
    Write-Info "Registering add-in via RegAsm: $dllPath"
    $p = Start-Process -FilePath $regasm -ArgumentList @('"' + $dllPath + '"', '/codebase', '/silent') -NoNewWindow -PassThru -Wait
    if ($p.ExitCode -ne 0) { throw "RegAsm failed with exit code $($p.ExitCode)" }
    Write-Info "Registration complete."
}

function Get-SwApp([int]$timeoutSeconds = 90) {
    Write-Info "Attempting to get SolidWorks application object..."
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($stopwatch.Elapsed.TotalSeconds -lt $timeoutSeconds) {
        try {
            $sw = [System.Runtime.InteropServices.Marshal]::GetActiveObject("SldWorks.Application")
            if ($sw -and $sw.Visible -ne $null) { # Check a property to ensure it's responsive
                Write-Info "SolidWorks application object acquired."
                return $sw
            }
        }
        catch { /* Ignore errors while waiting */ }
        Start-Sleep -Seconds 2
    }
    throw "Could not get SolidWorks application object within $($timeoutSeconds)s timeout."
}

try {
    if (-not (Test-Path $Root)) { throw "Root not found: $Root" }

    # Discover files
    $files = Get-ChildItem -Path $Root -Recurse -File |
        Where-Object { $Extensions -contains $_.Extension.ToLowerInvariant() } |
        Sort-Object FullName
    if (-not $files) { throw "No test files found." }
    Write-Info "Found $($files.Count) files."

    # Resolve and register add-in
    if (-not $AddInPath) {
        $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
        $AddInPath = Join-Path $repoRoot 'bin\Debug\swcsharpaddin.dll'
    }
    Register-AddIn -dllPath $AddInPath

    # Clean up old processes and start fresh
    Write-Info "Ensuring no old SolidWorks processes are running..."
    Get-Process sldworks -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2

    Write-Info "Launching SolidWorks: $SwExePath"
    Start-Process -FilePath $SwExePath -ArgumentList "/b" # /b = start minimized without splash
    
    $sw = Get-SwApp
    $sw.Visible = $false

    # Get add-in object with retry
    $addin = $null
    $retries = 5
    while ($retries-- -gt 0 -and -not $addin) {
        try { $addin = $sw.GetAddInObject($AddInGuid) } catch {}
        if (-not $addin) { Start-Sleep -Seconds 1 }
    }
    if (-not $addin) { throw "GetAddInObject failed for GUID: $AddInGuid" }
    Write-Info "Add-in object acquired."

    $results = foreach ($f in $files) {
        Write-Info "Processing: $($f.Name)"
        try {
            $ok = $addin.Automation_RunConvertToSheetMetal($f.FullName)
            [pscustomobject]@{ File = $f.Name; Result = if ($ok) { 'OK' } else { 'FAIL' } }
        }
        catch {
            Write-Warning "Exception on $($f.Name): $($_.Exception.Message)"
            [pscustomobject]@{ File = $f.Name; Result = 'EXCEPTION' }
        }
    }

    # Write summary
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outDir = 'C:\SolidWorksMacroLogs'
    if (-not (Test-Path $outDir)) { New-Item -Path $outDir -ItemType Directory | Out-Null }
    $summaryPath = Join-Path $outDir ("AutomationSummary_$ts.csv")
    $results | Export-Csv -NoTypeInformation -Path $summaryPath

    $results | Format-Table
    Write-Host "`nDone. Summary: $summaryPath"
}
catch {
    Write-Error $_
}
finally {
    Write-Info "Closing SolidWorks..."
    try { if ($sw) { $sw.ExitApp() } } catch {}
    Start-Sleep -Seconds 5 # Give time to close gracefully
    Get-Process sldworks -ErrorAction SilentlyContinue | Stop-Process -Force
    if ($sw) { try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($sw) } catch {} }
}
