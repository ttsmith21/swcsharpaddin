param([string]$AddInPath = '.\bin\Debug\swcsharpaddin.dll')

Write-Host "Manual test: Launch SolidWorks normally, then run automation connection test"

if (-not (Test-Path $AddInPath)) {
    Write-Error "Add-in not found: $AddInPath"
    exit 1
}

# Register add-in
$regasm = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
Write-Host "Registering: $AddInPath"
& $regasm $AddInPath /codebase /silent

Write-Host "Please:"
Write-Host "1. Manually launch SolidWorks"
Write-Host "2. Ensure the C# Addin is loaded (Tools > Add-Ins)"
Write-Host "3. Press Enter to continue with automation test"
Read-Host

try {
    $sw = [System.Runtime.InteropServices.Marshal]::GetActiveObject("SldWorks.Application")
    Write-Host "? Found SolidWorks application"
    
    $addin = $sw.GetAddInObject('{d5355548-9569-4381-9939-5d14252a3e47}')
    if ($addin) {
        Write-Host "? Found add-in object"
        
        # Test one file
        $testFile = Get-ChildItem 'C:\SWTests' -Recurse -File | Where-Object { $_.Extension -in '.sldprt','.stp','.step' } | Select-Object -First 1
        if ($testFile) {
            Write-Host "Testing: $($testFile.FullName)"
            $result = $addin.Automation_RunConvertToSheetMetal($testFile.FullName)
            Write-Host "Result: $result"
        } else {
            Write-Host "No test files found in C:\SWTests"
        }
    } else {
        Write-Host "? Add-in object not found"
    }
} catch {
    Write-Host "? Error: $($_.Exception.Message)"
}

Write-Host "Done. Check C:\SolidWorksMacroLogs\ErrorLog.txt for details."