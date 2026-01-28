<#
.SYNOPSIS
    Analyzes build errors and suggests or auto-fixes common issues.
.DESCRIPTION
    Reads build-errors.txt, identifies error patterns, and:
    - CS0246/CS0103: Missing using directives (auto-fixable)
    - CS1061: Method not found - suggests correct method
    - CS0535: Interface not implemented - lists missing members
    - CS7036: Missing argument - shows expected signature
.PARAMETER AutoFix
    If specified, automatically fixes errors where possible.
.PARAMETER Verbose
    Show additional diagnostic information.
#>

param(
    [switch]$AutoFix,
    [switch]$Verbose
)

$projectRoot = Split-Path -Parent $PSScriptRoot
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
    "PerformanceTracker" = "NM.Core"

    # NM.Core.Models
    "SwModelInfo" = "NM.Core.Models"

    # NM.Core.Manufacturing
    "PartMetrics" = "NM.Core.Manufacturing"
    "CalcResult" = "NM.Core.Manufacturing"
    "CutMetrics" = "NM.Core.Manufacturing"
    "TotalCostInputs" = "NM.Core.Manufacturing"
    "TotalCostCalculator" = "NM.Core.Manufacturing"
    "F140Calculator" = "NM.Core.Manufacturing"
    "F220Calculator" = "NM.Core.Manufacturing"
    "F325Calculator" = "NM.Core.Manufacturing"

    # NM.Core.DataModel
    "PartData" = "NM.Core.DataModel"
    "ProcessingStatus" = "NM.Core.DataModel"

    # NM.Core.Processing
    "SimpleTubeProcessor" = "NM.Core.Processing"
    "TubeGeometry" = "NM.Core.Processing"
    "SimpleSheetMetalProcessor" = "NM.Core.Processing"

    # NM.Core.ProblemParts
    "ProblemPartManager" = "NM.Core.ProblemParts"
    "ProblemCategory" = "NM.Core.ProblemParts"

    # NM.Core.Validation
    "ValidationResult" = "NM.Core.Validation"
    "PartValidator" = "NM.Core.Validation"

    # NM.SwAddin types
    "PartValidationAdapter" = "NM.SwAddin.Validation"
    "PartPreflight" = "NM.SwAddin.Validation"
    "PreflightResult" = "NM.SwAddin.Validation"
    "MainRunner" = "NM.SwAddin.Pipeline"
    "FolderProcessor" = "NM.SwAddin.Processing"
    "GenericPartProcessor" = "NM.SwAddin.Processing"
    "TubePartProcessor" = "NM.SwAddin.Processing"
    "ProcessingCoordinator" = "NM.SwAddin.Processing"
    "CustomPropertiesService" = "NM.SwAddin.Properties"
    "BendAnalyzer" = "NM.SwAddin.Manufacturing"
    "FlatPatternAnalyzer" = "NM.SwAddin.Manufacturing"
    "ComponentCollector" = "NM.SwAddin.Assembly"
    "ComponentValidator" = "NM.SwAddin.Assembly"
    "AssemblyPreprocessor" = "NM.SwAddin.Assembly"
    "FaceAnalyzer" = "NM.SwAddin.Geometry"

    # SolidWorks types
    "IModelDoc2" = "SolidWorks.Interop.sldworks"
    "ISldWorks" = "SolidWorks.Interop.sldworks"
    "IPartDoc" = "SolidWorks.Interop.sldworks"
    "IAssemblyDoc" = "SolidWorks.Interop.sldworks"
    "IDrawingDoc" = "SolidWorks.Interop.sldworks"
    "IFace2" = "SolidWorks.Interop.sldworks"
    "IBody2" = "SolidWorks.Interop.sldworks"
    "IFeature" = "SolidWorks.Interop.sldworks"
    "IComponent2" = "SolidWorks.Interop.sldworks"
    "ISurface" = "SolidWorks.Interop.sldworks"
    "IEdge" = "SolidWorks.Interop.sldworks"
    "IVertex" = "SolidWorks.Interop.sldworks"
    "ILoop2" = "SolidWorks.Interop.sldworks"
    "IConfiguration" = "SolidWorks.Interop.sldworks"
    "IModelDocExtension" = "SolidWorks.Interop.sldworks"
    "ICustomPropertyManager" = "SolidWorks.Interop.sldworks"
    "IBomTableAnnotation" = "SolidWorks.Interop.sldworks"
    "swDocumentTypes_e" = "SolidWorks.Interop.swconst"
    "swBodyType_e" = "SolidWorks.Interop.swconst"
    "swComponentSuppressionState_e" = "SolidWorks.Interop.swconst"
    "swBomType_e" = "SolidWorks.Interop.swconst"
    "swNumberingType_e" = "SolidWorks.Interop.swconst"
    "ISwAddin" = "SolidWorks.Interop.swpublished"

    # System types
    "StringBuilder" = "System.Text"
    "Regex" = "System.Text.RegularExpressions"
    "Path" = "System.IO"
    "File" = "System.IO"
    "Directory" = "System.IO"
    "Debug" = "System.Diagnostics"
}

# Common method corrections (CS1061)
$methodCorrections = @{
    # SolidWorks API corrections
    "GetCylinderParams2" = "CylinderParams (property, not method)"
    "GetPlaneParams2" = "PlaneParams (property, not method)"
    "GetVertices" = "IGetStartVertex() / IGetEndVertex() for IEdge"
    "GetBodies" = "GetBodies2(int bodyType, bool visibleOnly)"
    "GetFaces2" = "GetFaces() - no parameters needed"
    "Save" = "Save3(int options, ref int errors, ref int warnings)"
    "OpenDoc" = "OpenDoc6(string fileName, int type, int options, string config, ref int errors, ref int warnings)"
    "InsertBomTable" = "InsertBomTable3(...) - check parameter count"
}

# Parse errors
$errors = Get-Content $errorLog
$usingFixes = @{}
$methodErrors = @()
$interfaceErrors = @()
$argumentErrors = @()
$otherErrors = @()

foreach ($line in $errors) {
    # CS0246/CS0103: Missing type/namespace
    if ($line -match "^(.+?)\((\d+),\d+\):\s*error CS0(246|103):.+?'(\w+)'") {
        $file = $matches[1]
        $lineNum = $matches[2]
        $typeName = $matches[4]

        if ($typeToNamespace.ContainsKey($typeName)) {
            $ns = $typeToNamespace[$typeName]
            if (-not $usingFixes.ContainsKey($file)) {
                $usingFixes[$file] = @{}
            }
            $usingFixes[$file][$ns] = $true
        } else {
            $otherErrors += @{
                Type = "CS0246/CS0103"
                File = $file
                Line = $lineNum
                Message = "Unknown type: $typeName"
                Suggestion = "Add type to fix-errors.ps1 typeToNamespace mapping"
            }
        }
    }
    # CS1061: Method not found
    elseif ($line -match "error CS1061:.+?'(\w+)'.+?'(\w+)'") {
        $typeName = $matches[1]
        $methodName = $matches[2]
        $suggestion = if ($methodCorrections.ContainsKey($methodName)) { $methodCorrections[$methodName] } else { "Check API documentation" }

        $methodErrors += @{
            File = if ($line -match "^(.+?)\(") { $matches[1] } else { "Unknown" }
            Line = if ($line -match "\((\d+),") { $matches[1] } else { "?" }
            Type = $typeName
            Method = $methodName
            Suggestion = $suggestion
        }
    }
    # CS0535: Interface not implemented
    elseif ($line -match "error CS0535:.+?'(.+?)'.+?'(.+?)'") {
        $className = $matches[1]
        $memberName = $matches[2]

        $interfaceErrors += @{
            File = if ($line -match "^(.+?)\(") { $matches[1] } else { "Unknown" }
            Class = $className
            MissingMember = $memberName
        }
    }
    # CS7036: Missing required argument
    elseif ($line -match "error CS7036:.+?'(\w+)'.+?'(.+?)'") {
        $paramName = $matches[1]
        $methodSig = $matches[2]

        $argumentErrors += @{
            File = if ($line -match "^(.+?)\(") { $matches[1] } else { "Unknown" }
            Line = if ($line -match "\((\d+),") { $matches[1] } else { "?" }
            Parameter = $paramName
            Method = $methodSig
        }
    }
}

# === Output Results ===

$totalFixable = $usingFixes.Count
$totalUnfixable = $methodErrors.Count + $interfaceErrors.Count + $argumentErrors.Count + $otherErrors.Count

Write-Host ""
Write-Host "=== Error Analysis ===" -ForegroundColor Cyan
Write-Host "Fixable (using directives): $totalFixable files" -ForegroundColor $(if ($totalFixable -gt 0) { "Yellow" } else { "Green" })
Write-Host "Manual fixes needed: $totalUnfixable issues" -ForegroundColor $(if ($totalUnfixable -gt 0) { "Yellow" } else { "Green" })

# === CS0246/CS0103: Missing Using Directives ===
if ($usingFixes.Count -gt 0) {
    Write-Host ""
    Write-Host "--- Missing Using Directives (Auto-fixable) ---" -ForegroundColor Yellow

    foreach ($file in $usingFixes.Keys) {
        $namespaces = $usingFixes[$file].Keys | Sort-Object
        $shortFile = $file -replace [regex]::Escape($projectRoot), "."
        Write-Host "`n$shortFile" -ForegroundColor White
        foreach ($ns in $namespaces) {
            Write-Host "  using $ns;" -ForegroundColor Gray
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
}

# === CS1061: Method Not Found ===
if ($methodErrors.Count -gt 0) {
    Write-Host ""
    Write-Host "--- Method Not Found (CS1061) - Manual Fix Required ---" -ForegroundColor Yellow

    foreach ($err in $methodErrors) {
        $shortFile = $err.File -replace [regex]::Escape($projectRoot), "."
        Write-Host "`n$shortFile`:$($err.Line)" -ForegroundColor White
        Write-Host "  Type: $($err.Type)" -ForegroundColor Gray
        Write-Host "  Missing: $($err.Method)" -ForegroundColor Red
        Write-Host "  Suggestion: $($err.Suggestion)" -ForegroundColor Cyan
    }
}

# === CS0535: Interface Not Implemented ===
if ($interfaceErrors.Count -gt 0) {
    Write-Host ""
    Write-Host "--- Interface Not Implemented (CS0535) - Manual Fix Required ---" -ForegroundColor Yellow

    $grouped = $interfaceErrors | Group-Object -Property Class
    foreach ($group in $grouped) {
        Write-Host "`nClass: $($group.Name)" -ForegroundColor White
        Write-Host "  Missing members:" -ForegroundColor Gray
        foreach ($err in $group.Group) {
            Write-Host "    - $($err.MissingMember)" -ForegroundColor Red
        }
    }
}

# === CS7036: Missing Argument ===
if ($argumentErrors.Count -gt 0) {
    Write-Host ""
    Write-Host "--- Missing Argument (CS7036) - Manual Fix Required ---" -ForegroundColor Yellow

    foreach ($err in $argumentErrors) {
        $shortFile = $err.File -replace [regex]::Escape($projectRoot), "."
        Write-Host "`n$shortFile`:$($err.Line)" -ForegroundColor White
        Write-Host "  Missing parameter: $($err.Parameter)" -ForegroundColor Red
        Write-Host "  Method signature: $($err.Method)" -ForegroundColor Gray
    }
}

# === Other Errors ===
if ($otherErrors.Count -gt 0 -and $Verbose) {
    Write-Host ""
    Write-Host "--- Other Errors ---" -ForegroundColor Yellow

    foreach ($err in $otherErrors) {
        $shortFile = $err.File -replace [regex]::Escape($projectRoot), "."
        Write-Host "$shortFile`:$($err.Line) - $($err.Message)" -ForegroundColor Gray
        Write-Host "  $($err.Suggestion)" -ForegroundColor DarkGray
    }
}

Write-Host ""
