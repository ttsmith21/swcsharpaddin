<#
.SYNOPSIS
    Analyzes CS0246/CS0103 errors and suggests or auto-fixes missing using directives.
.DESCRIPTION
    Reads build-errors.txt, identifies missing type errors, and suggests the correct
    using directive based on a known mapping. Can optionally auto-fix files.
.PARAMETER AutoFix
    If specified, automatically adds missing using directives to files.
#>

param(
    [switch]$AutoFix
)

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not $projectRoot) { $projectRoot = (Get-Location).Path }

$errorLog = Join-Path $projectRoot "build-errors.txt"

if (-not (Test-Path $errorLog)) {
    Write-Host "No build-errors.txt found. Run build-quick.ps1 first." -ForegroundColor Yellow
    exit 0
}

# Known type -> namespace mappings for this project
$typeToNamespace = @{
    # NM.Core types
    "ErrorHandler" = "NM.Core"
    "ModelInfo" = "NM.Core"
    "ProcessingOptions" = "NM.Core"
    "CustomPropertyData" = "NM.Core"
    "CustomPropertyType" = "NM.Core"
    "Configuration" = "NM.Core"
    "DifficultyLevel" = "NM.Core"

    # NM.Core.Models
    "SwModelInfo" = "NM.Core.Models"

    # NM.Core.Manufacturing
    "PartMetrics" = "NM.Core.Manufacturing"
    "CalcResult" = "NM.Core.Manufacturing"
    "CutMetrics" = "NM.Core.Manufacturing"
    "TotalCostInputs" = "NM.Core.Manufacturing"
    "TotalCostCalculator" = "NM.Core.Manufacturing"

    # NM.Core.DataModel
    "PartData" = "NM.Core.DataModel"
    "ProcessingStatus" = "NM.Core.DataModel"

    # NM.Core.Processing
    "SimpleTubeProcessor" = "NM.Core.Processing"
    "TubeGeometry" = "NM.Core.Processing"

    # NM.Core.ProblemParts
    "ProblemPartManager" = "NM.Core.ProblemParts"

    # NM.Core.Validation
    "ValidationResult" = "NM.Core.Validation"
    "PartValidator" = "NM.Core.Validation"

    # NM.SwAddin types
    "PartValidationAdapter" = "NM.SwAddin.Validation"
    "PartPreflight" = "NM.SwAddin.Validation"
    "MainRunner" = "NM.SwAddin"
    "FolderProcessor" = "NM.SwAddin.Processing"

    # SolidWorks types
    "IModelDoc2" = "SolidWorks.Interop.sldworks"
    "ISldWorks" = "SolidWorks.Interop.sldworks"
    "IPartDoc" = "SolidWorks.Interop.sldworks"
    "IAssemblyDoc" = "SolidWorks.Interop.sldworks"
    "IFace2" = "SolidWorks.Interop.sldworks"
    "IBody2" = "SolidWorks.Interop.sldworks"
    "IFeature" = "SolidWorks.Interop.sldworks"
    "IComponent2" = "SolidWorks.Interop.sldworks"
    "swDocumentTypes_e" = "SolidWorks.Interop.swconst"
    "swBodyType_e" = "SolidWorks.Interop.swconst"
    "ISwAddin" = "SolidWorks.Interop.swpublished"
}

# Parse errors
$errors = Get-Content $errorLog
$fixes = @{}

foreach ($line in $errors) {
    # Match: file(line,col): error CS0246: The type or namespace name 'TypeName' could not be found
    # Or: error CS0103: The name 'Name' does not exist
    if ($line -match "^(.+?)\((\d+),\d+\):\s*error CS0(246|103):.+?'(\w+)'") {
        $file = $matches[1]
        $lineNum = $matches[2]
        $typeName = $matches[4]

        if ($typeToNamespace.ContainsKey($typeName)) {
            $ns = $typeToNamespace[$typeName]
            if (-not $fixes.ContainsKey($file)) {
                $fixes[$file] = @{}
            }
            $fixes[$file][$ns] = $true
        } else {
            Write-Host "Unknown type: $typeName (in $file)" -ForegroundColor DarkGray
        }
    }
}

if ($fixes.Count -eq 0) {
    Write-Host "No fixable missing using directives found." -ForegroundColor Green
    exit 0
}

Write-Host "`n=== Missing Using Directives ===" -ForegroundColor Cyan

foreach ($file in $fixes.Keys) {
    $namespaces = $fixes[$file].Keys | Sort-Object
    Write-Host "`n$file" -ForegroundColor Yellow
    foreach ($ns in $namespaces) {
        Write-Host "  using $ns;" -ForegroundColor White
    }

    if ($AutoFix) {
        # Read file content
        $content = Get-Content $file -Raw
        $insertPoint = 0

        # Find last using statement
        if ($content -match "(?s)^(.*using [^;]+;\r?\n)") {
            $lastUsing = [regex]::Matches($content, "using [^;]+;\r?\n") | Select-Object -Last 1
            $insertPoint = $lastUsing.Index + $lastUsing.Length
        }

        # Build new usings to add
        $existingUsings = [regex]::Matches($content, "using ([^;]+);") | ForEach-Object { $_.Groups[1].Value }
        $newUsings = ""
        foreach ($ns in $namespaces) {
            if ($existingUsings -notcontains $ns) {
                $newUsings += "using $ns;`r`n"
            }
        }

        if ($newUsings) {
            $newContent = $content.Insert($insertPoint, $newUsings)
            Set-Content -Path $file -Value $newContent -NoNewline
            Write-Host "  [FIXED]" -ForegroundColor Green
        }
    }
}

if (-not $AutoFix) {
    Write-Host "`nRun with -AutoFix to automatically add these using directives." -ForegroundColor Cyan
}
