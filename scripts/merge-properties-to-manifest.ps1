<#
.SYNOPSIS
    Merges VBA custom properties from vba_properties.json into manifest-v2.json.

.DESCRIPTION
    Reads property dump from ReadPropertiesOnly QA mode and updates manifest-v2.json
    vbaBaseline sections with additional fields (costs, flat area, cut length, etc.)
    that the VBA macro wrote as custom properties.

.PARAMETER PropertiesPath
    Path to vba_properties.json. Auto-detects from latest Run_*.

.PARAMETER ManifestPath
    Path to manifest-v2.json. Default: tests/GoldStandard_Baseline/manifest-v2.json

.PARAMETER DryRun
    Preview changes without writing manifest.
#>

param(
    [string]$PropertiesPath = "",
    [string]$ManifestPath = "",
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$script:RepoRoot = Split-Path -Parent $PSScriptRoot
if (-not $script:RepoRoot -or -not (Test-Path $script:RepoRoot)) {
    $script:RepoRoot = Get-Location
}

function Write-Status($msg, $color = "White") { Write-Host $msg -ForegroundColor $color }

# ============ PROPERTY MAPPINGS ============
# VBA custom property name -> manifest vbaBaseline field name
# Format: @{ VbaName = @{ ManifestField = "name"; Convert = "none"|"minToHours"; IsNumeric = $true|$false } }

$PropertyMap = @{
    # Weight & material
    "RawWeight"     = @{ Field = "rawWeight"; Convert = "none"; Numeric = $true }
    "OptiMaterial"  = @{ Field = "optiMaterial"; Convert = "none"; Numeric = $false }
    "Description"   = @{ Field = "description"; Convert = "none"; Numeric = $false }
    "Thickness"     = @{ Field = "expectedThickness_in"; Convert = "none"; Numeric = $true }
    "Material"      = @{ Field = "material"; Convert = "none"; Numeric = $false }

    # Costs
    "MaterialCost"  = @{ Field = "materialCost"; Convert = "none"; Numeric = $true }
    "TotalCost"     = @{ Field = "totalCost"; Convert = "none"; Numeric = $true }
    "F115_Price"    = @{ Field = "laserCost"; Convert = "none"; Numeric = $true }
    "F140_Price"    = @{ Field = "bendCost"; Convert = "none"; Numeric = $true }
    "F210_Price"    = @{ Field = "deburCost"; Convert = "none"; Numeric = $true }
    "F220_Price"    = @{ Field = "tapCost"; Convert = "none"; Numeric = $true }
    "F325_Price"    = @{ Field = "rollCost"; Convert = "none"; Numeric = $true }

    # Sheet metal geometry
    "BendCount"     = @{ Field = "expectedBendCount"; Convert = "none"; Numeric = $true }
    "FlatArea"      = @{ Field = "flatArea_sqin"; Convert = "none"; Numeric = $true }
    "CutLength"     = @{ Field = "cutLength_in"; Convert = "none"; Numeric = $true }
    "BBoxLength"    = @{ Field = "bboxLength_in"; Convert = "none"; Numeric = $true }
    "BBoxWidth"     = @{ Field = "bboxWidth_in"; Convert = "none"; Numeric = $true }

    # Classification
    "IsSheetMetal"  = @{ Field = "isSheetMetal"; Convert = "none"; Numeric = $false }
    "IsTube"        = @{ Field = "isTube"; Convert = "none"; Numeric = $false }

    # Tube geometry
    "TubeOD"        = @{ Field = "tubeOD_in"; Convert = "none"; Numeric = $true }
    "TubeWall"      = @{ Field = "tubeWall_in"; Convert = "none"; Numeric = $true }
    "TubeLength"    = @{ Field = "tubeLength_in"; Convert = "none"; Numeric = $true }
    "TubeID"        = @{ Field = "tubeID_in"; Convert = "none"; Numeric = $true }
}

# Routing property mappings: VBA property -> (workCenter, field, convert)
$RoutingMap = @{
    "OP20_S"  = @{ WC = "O120"; SubField = "setup"; Convert = "minToHours" }
    "OP20_R"  = @{ WC = "O120"; SubField = "run"; Convert = "minToHours" }
    "F140_S"  = @{ WC = "N120"; SubField = "setup"; Convert = "minToHours" }
    "F140_R"  = @{ WC = "N120"; SubField = "run"; Convert = "minToHours" }
    "F210_S"  = @{ WC = "N140"; SubField = "setup"; Convert = "minToHours" }
    "F210_R"  = @{ WC = "N140"; SubField = "run"; Convert = "minToHours" }
    "F220_S"  = @{ WC = "N220"; SubField = "setup"; Convert = "minToHours" }
    "F220_R"  = @{ WC = "N220"; SubField = "run"; Convert = "minToHours" }
    "F325_S"  = @{ WC = "F325"; SubField = "setup"; Convert = "minToHours" }
    "F325_R"  = @{ WC = "F325"; SubField = "run"; Convert = "minToHours" }
}

# ============ MAIN ============

Write-Status ""
Write-Status "============================================" "Cyan"
Write-Status " MERGE VBA PROPERTIES TO MANIFEST" "Cyan"
Write-Status "============================================" "Cyan"
Write-Status ""

# Auto-detect properties file
if (-not $PropertiesPath) {
    $latestDir = Join-Path $script:RepoRoot "tests\Run_Latest"
    $propsFile = Join-Path $latestDir "vba_properties.json"
    if (Test-Path $propsFile) {
        $PropertiesPath = $propsFile
    } else {
        $runsDir = Join-Path $script:RepoRoot "tests"
        $latestRun = Get-ChildItem $runsDir -Directory -Filter "Run_*" -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending | Select-Object -First 1
        if ($latestRun) {
            $PropertiesPath = Join-Path $latestRun.FullName "vba_properties.json"
        }
    }
}

# Default manifest path
if (-not $ManifestPath) {
    $ManifestPath = Join-Path $script:RepoRoot "tests\GoldStandard_Baseline\manifest-v2.json"
}

if (-not $PropertiesPath -or -not (Test-Path $PropertiesPath)) {
    Write-Status "ERROR: vba_properties.json not found: $PropertiesPath" "Red"
    Write-Status "Run QA with Mode=ReadPropertiesOnly first." "Yellow"
    exit 1
}
if (-not (Test-Path $ManifestPath)) {
    Write-Status "ERROR: manifest-v2.json not found: $ManifestPath" "Red"
    exit 1
}

Write-Status "Properties: $PropertiesPath"
Write-Status "Manifest:   $ManifestPath"
if ($DryRun) { Write-Status "Mode:       DRY RUN (no changes written)" "Yellow" }
Write-Status ""

try {
    $propsData = Get-Content $PropertiesPath -Raw | ConvertFrom-Json
    $manifest = Get-Content $ManifestPath -Raw | ConvertFrom-Json
} catch {
    Write-Status "ERROR: Failed to parse JSON: $_" "Red"
    exit 1
}

$totalUpdated = 0
$totalSkipped = 0
$totalNewFields = 0

foreach ($fileProp in $propsData.files.PSObject.Properties) {
    $fileName = $fileProp.Name
    $props = $fileProp.Value

    # Find matching manifest entry (case-insensitive)
    $manifestEntry = $null
    foreach ($mProp in $manifest.files.PSObject.Properties) {
        if ($mProp.Name -ieq $fileName) {
            $manifestEntry = $mProp.Value
            break
        }
    }

    if (-not $manifestEntry) {
        Write-Status "  SKIP: $fileName (not in manifest)" "DarkGray"
        $totalSkipped++
        continue
    }

    # Ensure vbaBaseline exists
    if (-not $manifestEntry.vbaBaseline) {
        $manifestEntry | Add-Member -NotePropertyName "vbaBaseline" -NotePropertyValue (New-Object PSObject) -Force
    }
    $baseline = $manifestEntry.vbaBaseline

    $fileUpdates = 0

    # Map simple properties
    foreach ($propName in $props.PSObject.Properties.Name) {
        if ($PropertyMap.ContainsKey($propName)) {
            $mapping = $PropertyMap[$propName]
            $value = $props.$propName

            # Skip empty values
            if ([string]::IsNullOrWhiteSpace($value)) { continue }

            # Convert numeric values
            $finalValue = $value
            if ($mapping.Numeric) {
                try {
                    $numVal = [double]$value
                    if ($numVal -eq 0) { continue } # Skip zero values
                    $finalValue = $numVal
                } catch { continue }
            }

            # Add/update field on vbaBaseline
            $fieldName = $mapping.Field
            if ($baseline.PSObject.Properties.Name -contains $fieldName) {
                $baseline.$fieldName = $finalValue
            } else {
                $baseline | Add-Member -NotePropertyName $fieldName -NotePropertyValue $finalValue -Force
                $totalNewFields++
            }
            $fileUpdates++
        }

        # Map routing properties
        if ($RoutingMap.ContainsKey($propName)) {
            $rMapping = $RoutingMap[$propName]
            $value = $props.$propName

            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            try { $numVal = [double]$value } catch { continue }
            if ($numVal -eq 0) { continue }

            # Convert minutes to hours
            if ($rMapping.Convert -eq "minToHours") {
                $numVal = $numVal / 60.0
            }

            # Ensure routing object exists
            if (-not $baseline.routing) {
                $baseline | Add-Member -NotePropertyName "routing" -NotePropertyValue (New-Object PSObject) -Force
            }
            $routing = $baseline.routing

            # Ensure work center exists
            $wc = $rMapping.WC
            if (-not ($routing.PSObject.Properties.Name -contains $wc)) {
                $routing | Add-Member -NotePropertyName $wc -NotePropertyValue (New-Object PSObject) -Force
            }
            $wcObj = $routing.$wc

            # Set the sub-field
            $subField = $rMapping.SubField
            if ($wcObj.PSObject.Properties.Name -contains $subField) {
                $wcObj.$subField = [Math]::Round($numVal, 6)
            } else {
                $wcObj | Add-Member -NotePropertyName $subField -NotePropertyValue ([Math]::Round($numVal, 6)) -Force
                $totalNewFields++
            }
            $fileUpdates++
        }
    }

    if ($fileUpdates -gt 0) {
        $totalUpdated++
        Write-Status "  $fileName`: $fileUpdates fields merged" "Green"
    } else {
        Write-Status "  $fileName`: no new data" "DarkGray"
    }
}

Write-Status ""
Write-Status "============================================" "Cyan"
Write-Status " SUMMARY" "Cyan"
Write-Status "============================================" "Cyan"
Write-Status ""
Write-Status "  Files updated:  $totalUpdated" "Green"
Write-Status "  Files skipped:  $totalSkipped" "DarkGray"
Write-Status "  New fields:     $totalNewFields" "Cyan"

if (-not $DryRun -and $totalUpdated -gt 0) {
    # Write updated manifest (with pretty formatting)
    $jsonOutput = $manifest | ConvertTo-Json -Depth 10
    $jsonOutput | Out-File $ManifestPath -Encoding UTF8
    Write-Status ""
    Write-Status "  Wrote updated manifest to: $ManifestPath" "Green"
} elseif ($DryRun) {
    Write-Status ""
    Write-Status "  DRY RUN: No changes written. Remove -DryRun to apply." "Yellow"
}

Write-Status ""
