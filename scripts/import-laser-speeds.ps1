<#
.SYNOPSIS
    Import per-alloy laser cutting speeds from Laser2022v4.xlsx into nm-tables.json.

.DESCRIPTION
    Reads each material tab (304L, 316L, 309, 2205, CS, AL) from the Excel workbook
    and imports them as byMaterial entries in nm-tables.json. Also refreshes the
    3 group tables (stainlessSteel, carbonSteel, aluminum) from matching sheets.

    This achieves VBA parity: the VBA macro read per-alloy tabs from this same Excel file.
    The C# pipeline previously used only 3 groups, losing per-alloy fidelity.

.PARAMETER ExcelPath
    Path to the Laser2022v4.xlsx workbook.

.PARAMETER JsonPath
    Path to nm-tables.json to update.

.PARAMETER DryRun
    Print what would change without writing to disk.

.EXAMPLE
    .\scripts\import-laser-speeds.ps1 -DryRun
    .\scripts\import-laser-speeds.ps1
    .\scripts\import-laser-speeds.ps1 -ExcelPath "C:\local\Laser2022v4.xlsx"
#>
param(
    [string]$ExcelPath = "O:\Engineering Department\Solidworks\Macros\(Semi)Autopilot\Laser2022v4.xlsx",
    [string]$JsonPath = "$PSScriptRoot\..\config\nm-tables.json",
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

# Resolve paths (PS 5.1 compatible)
$resolved = Resolve-Path $JsonPath -ErrorAction SilentlyContinue
if ($resolved) { $JsonPath = $resolved.Path }

if (-not (Test-Path $ExcelPath)) {
    Write-Error "Excel file not found: $ExcelPath`nProvide the path via -ExcelPath parameter."
    exit 1
}

if (-not (Test-Path $JsonPath)) {
    Write-Error "nm-tables.json not found: $JsonPath"
    exit 1
}

# Excel tab -> JSON key mapping for per-material overrides
$ByMaterialTabs = @{
    '304L' = '304L'
    '316L' = '316L'
    '309'  = '309'
    '2205' = '2205'
    'CS'   = 'CS'
    'AL'   = 'AL'
}

# Excel tab -> JSON group key mapping (refresh existing groups)
$GroupTabs = @{
    'Stainless Steel' = 'stainlessSteel'
    'Carbon Steel'    = 'carbonSteel'
    'Aluminum'        = 'aluminum'
}

Write-Host "=== Laser Speed Import ===" -ForegroundColor Cyan
Write-Host "Excel: $ExcelPath"
Write-Host "JSON:  $JsonPath"
if ($DryRun) { Write-Host "[DRY RUN] No files will be modified." -ForegroundColor Yellow }

function Read-SpeedSheet {
    param(
        [object]$Worksheet,
        [string]$Label
    )
    <#
    Expected layout (VBA convention):
      Column B: Thickness (inches)
      Column C: Feed rate (IPM)
      Column D: Pierce time (seconds)
      Row 1: Headers
      Data starts at row 2
    #>
    $entries = @()
    $row = 2
    $maxEmpty = 3  # Stop after 3 consecutive empty rows
    $emptyCount = 0

    while ($emptyCount -lt $maxEmpty) {
        $thickness = $Worksheet.Cells.Item($row, 2).Value2  # Column B
        $feedRate  = $Worksheet.Cells.Item($row, 3).Value2  # Column C
        $pierce    = $Worksheet.Cells.Item($row, 4).Value2  # Column D

        if ($null -eq $thickness -or $thickness -eq 0) {
            $emptyCount++
            $row++
            continue
        }
        $emptyCount = 0

        # Validate numeric values
        $t = [double]$thickness
        $f = 0.0
        if ($null -ne $feedRate -and $feedRate -ne '') { $f = [double]$feedRate }
        $p = 0.0
        if ($null -ne $pierce -and $pierce -ne '') { $p = [double]$pierce }

        if ($t -gt 0) {
            # Keep sentinel entries (feedRate=0) as "can't cut" markers
            $entries += [PSCustomObject]@{
                thicknessIn   = [Math]::Round($t, 4)
                feedRateIpm   = [Math]::Round($f, 4)
                pierceSeconds = [Math]::Round($p, 4)
            }
        }
        $row++
    }

    # Sort ascending by thickness (lookup algorithm expects thinnest first)
    $entries = @($entries | Sort-Object { $_.thicknessIn })

    $color = 'Green'
    if ($entries.Count -eq 0) { $color = 'Red' }
    Write-Host "  $Label : $($entries.Count) entries (rows 2-$row)" -ForegroundColor $color
    return ,$entries
}

# Open Excel
Write-Host ""
Write-Host "Opening Excel..." -ForegroundColor Gray
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($ExcelPath, 0, $true)  # ReadOnly

    # Get available sheet names
    $sheetNames = @()
    foreach ($ws in $workbook.Worksheets) {
        $sheetNames += $ws.Name
    }
    Write-Host "Found sheets: $($sheetNames -join ', ')" -ForegroundColor Gray

    # Read JSON
    $json = Get-Content $JsonPath -Raw | ConvertFrom-Json

    # Ensure laserSpeeds exists
    if (-not $json.laserSpeeds) {
        Write-Error "nm-tables.json does not contain 'laserSpeeds' section"
        exit 1
    }

    # === Import per-material tabs into byMaterial ===
    Write-Host ""
    Write-Host "--- Per-Material Tabs (byMaterial) ---" -ForegroundColor Cyan

    # Initialize byMaterial as a PSCustomObject if not present
    if (-not $json.laserSpeeds.byMaterial) {
        $json.laserSpeeds | Add-Member -NotePropertyName 'byMaterial' -NotePropertyValue ([PSCustomObject]@{}) -Force
    }

    $importedCount = 0
    foreach ($tab in $ByMaterialTabs.GetEnumerator()) {
        $sheetName = $tab.Key
        $jsonKey = $tab.Value

        if ($sheetNames -notcontains $sheetName) {
            Write-Host "  SKIP: Sheet '$sheetName' not found in workbook" -ForegroundColor Yellow
            continue
        }

        $ws = $workbook.Worksheets.Item($sheetName)
        $entries = Read-SpeedSheet -Worksheet $ws -Label $sheetName

        if ($entries.Count -gt 0) {
            $json.laserSpeeds.byMaterial | Add-Member -NotePropertyName $jsonKey -NotePropertyValue $entries -Force
            $importedCount++
        }
    }
    Write-Host "Imported $importedCount per-material tables into byMaterial." -ForegroundColor Cyan

    # === Refresh group tables ===
    Write-Host ""
    Write-Host "--- Group Tables (refresh) ---" -ForegroundColor Cyan

    foreach ($tab in $GroupTabs.GetEnumerator()) {
        $sheetName = $tab.Key
        $jsonKey = $tab.Value

        if ($sheetNames -notcontains $sheetName) {
            Write-Host "  SKIP: Sheet '$sheetName' not found - keeping existing $jsonKey" -ForegroundColor Yellow
            continue
        }

        $ws = $workbook.Worksheets.Item($sheetName)
        $entries = Read-SpeedSheet -Worksheet $ws -Label "$sheetName -> $jsonKey"

        if ($entries.Count -gt 0) {
            $json.laserSpeeds.$jsonKey = $entries
        }
    }

    # === Summary ===
    Write-Host ""
    Write-Host "--- Summary ---" -ForegroundColor Cyan
    $byMat = $json.laserSpeeds.byMaterial
    foreach ($prop in $byMat.PSObject.Properties) {
        Write-Host "  byMaterial.$($prop.Name): $($prop.Value.Count) entries"
    }
    foreach ($group in @('stainlessSteel', 'carbonSteel', 'aluminum')) {
        $g = $json.laserSpeeds.$group
        $count = 0
        if ($g) { $count = $g.Count }
        Write-Host "  ${group}: $count entries"
    }

    # Write JSON
    if (-not $DryRun) {
        $jsonOut = $json | ConvertTo-Json -Depth 10
        Set-Content -Path $JsonPath -Value $jsonOut -Encoding UTF8
        Write-Host ""
        Write-Host "Wrote updated nm-tables.json to: $JsonPath" -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "[DRY RUN] Would write updated nm-tables.json to: $JsonPath" -ForegroundColor Yellow
        # Show a sample of what byMaterial looks like
        $sample = $json.laserSpeeds.byMaterial | ConvertTo-Json -Depth 5
        $lines = $sample -split "`n"
        if ($lines.Count -gt 30) {
            $preview = $lines[0..29] -join "`n"
            $remaining = $lines.Count - 30
            Write-Host $preview
            Write-Host "... ($remaining more lines)"
        } else {
            Write-Host $sample
        }
    }

} finally {
    # Clean up COM objects
    if ($workbook) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Host ""
    Write-Host "Excel closed." -ForegroundColor Gray
}

Write-Host ""
Write-Host "Done. Next steps:" -ForegroundColor Cyan
Write-Host "  1. Review diff: git diff config/nm-tables.json"
Write-Host "  2. Build: .\scripts\build-and-test.ps1 -SkipClean"
Write-Host "  3. Test: dotnet test src/NM.Core.Tests --filter FullyQualifiedName~RealSpeedTable --verbosity normal"
