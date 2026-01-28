<#
.SYNOPSIS
    Pre-build validation to catch common issues before MSBuild runs.
.DESCRIPTION
    Validates:
    1. All .cs files in src/ are included in .csproj
    2. Using directives reference known namespaces
    3. No obvious syntax errors (unmatched braces)
    4. No duplicate class definitions
.PARAMETER Fix
    If specified, automatically fix issues where possible.
#>

param(
    [switch]$Fix
)

$projectRoot = Split-Path -Parent $PSScriptRoot
$csprojPath = Join-Path $projectRoot "swcsharpaddin.csproj"

Write-Host "=== Pre-Build Validation ===" -ForegroundColor Cyan
Write-Host "Project: $csprojPath"
Write-Host ""

$issues = @()
$warnings = @()

# === 1. Check all .cs files are in .csproj ===
Write-Host "Checking .cs files in .csproj..." -ForegroundColor Yellow

$csprojContent = Get-Content $csprojPath -Raw
$csFiles = Get-ChildItem -Path $projectRoot -Filter "*.cs" -Recurse |
    Where-Object { $_.FullName -notmatch "\\(obj|bin|packages|\.vs)\\" } |
    ForEach-Object { $_.FullName }

$missingFromCsproj = @()
foreach ($file in $csFiles) {
    $relativePath = $file -replace [regex]::Escape("$projectRoot\"), ""
    # Check various Include patterns
    $patterns = @(
        [regex]::Escape("Include=`"$relativePath`""),
        [regex]::Escape("Include=`".\$relativePath`""),
        [regex]::Escape($relativePath)
    )

    $found = $false
    foreach ($pattern in $patterns) {
        if ($csprojContent -match $pattern) {
            $found = $true
            break
        }
    }

    if (-not $found) {
        $missingFromCsproj += $relativePath
    }
}

if ($missingFromCsproj.Count -gt 0) {
    Write-Host "  Missing from .csproj: $($missingFromCsproj.Count) files" -ForegroundColor Red
    foreach ($file in $missingFromCsproj | Select-Object -First 10) {
        Write-Host "    - $file" -ForegroundColor Gray
    }
    if ($missingFromCsproj.Count -gt 10) {
        Write-Host "    ... and $($missingFromCsproj.Count - 10) more" -ForegroundColor Gray
    }
    $issues += "Missing .cs files in .csproj"

    if ($Fix) {
        Write-Host "  Running sync-csproj.ps1..." -ForegroundColor Cyan
        & "$PSScriptRoot\sync-csproj.ps1"
    } else {
        Write-Host "  Run: .\scripts\sync-csproj.ps1" -ForegroundColor Cyan
    }
} else {
    Write-Host "  All .cs files included in .csproj" -ForegroundColor Green
}

# === 2. Check for unmatched braces ===
Write-Host ""
Write-Host "Checking for unmatched braces..." -ForegroundColor Yellow

$braceIssues = @()
foreach ($file in $csFiles) {
    $content = Get-Content $file -Raw -ErrorAction SilentlyContinue
    if ($null -eq $content) { continue }

    # Simple brace counting (ignores strings/comments, but catches obvious issues)
    $opens = ([regex]::Matches($content, '\{')).Count
    $closes = ([regex]::Matches($content, '\}')).Count

    if ($opens -ne $closes) {
        $diff = $opens - $closes
        $relativePath = $file -replace [regex]::Escape("$projectRoot\"), ""
        $braceIssues += @{
            File = $relativePath
            Open = $opens
            Close = $closes
            Diff = $diff
        }
    }
}

if ($braceIssues.Count -gt 0) {
    Write-Host "  Potential brace mismatches: $($braceIssues.Count) files" -ForegroundColor Red
    foreach ($issue in $braceIssues | Select-Object -First 5) {
        $sign = if ($issue.Diff -gt 0) { "+" } else { "" }
        Write-Host "    $($issue.File): $($sign)$($issue.Diff) braces" -ForegroundColor Gray
    }
    $warnings += "Potential brace mismatches"
} else {
    Write-Host "  No obvious brace mismatches" -ForegroundColor Green
}

# === 3. Check for duplicate class definitions ===
Write-Host ""
Write-Host "Checking for duplicate classes..." -ForegroundColor Yellow

$classDefinitions = @{}
foreach ($file in $csFiles) {
    $content = Get-Content $file -Raw -ErrorAction SilentlyContinue
    if ($null -eq $content) { continue }

    # Find class/struct/interface definitions
    $matches = [regex]::Matches($content, '(?:public|internal|private|protected)?\s*(?:sealed|abstract|static|partial)?\s*(?:class|struct|interface)\s+(\w+)')
    foreach ($match in $matches) {
        $className = $match.Groups[1].Value
        $relativePath = $file -replace [regex]::Escape("$projectRoot\"), ""

        if ($classDefinitions.ContainsKey($className)) {
            # Check if it's a partial class (allowed)
            if ($content -notmatch "partial\s+(?:class|struct)\s+$className") {
                if (-not $classDefinitions[$className].Contains($relativePath)) {
                    $classDefinitions[$className] += $relativePath
                }
            }
        } else {
            $classDefinitions[$className] = @($relativePath)
        }
    }
}

$duplicates = $classDefinitions.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 }
if ($duplicates) {
    Write-Host "  Potential duplicate classes:" -ForegroundColor Red
    foreach ($dup in $duplicates) {
        Write-Host "    $($dup.Key):" -ForegroundColor White
        foreach ($file in $dup.Value) {
            Write-Host "      - $file" -ForegroundColor Gray
        }
    }
    $issues += "Duplicate class definitions"
} else {
    Write-Host "  No duplicate classes found" -ForegroundColor Green
}

# === 4. Check for common using directive issues ===
Write-Host ""
Write-Host "Checking using directives..." -ForegroundColor Yellow

$knownNamespaces = @(
    "System", "System.Collections", "System.Collections.Generic", "System.IO",
    "System.Linq", "System.Text", "System.Text.RegularExpressions", "System.Diagnostics",
    "System.Runtime.InteropServices", "System.Windows.Forms", "System.Drawing",
    "NM.Core", "NM.Core.Models", "NM.Core.Manufacturing", "NM.Core.DataModel",
    "NM.Core.Processing", "NM.Core.ProblemParts", "NM.Core.Validation", "NM.Core.Materials",
    "NM.SwAddin", "NM.SwAddin.Validation", "NM.SwAddin.Processing", "NM.SwAddin.Pipeline",
    "NM.SwAddin.Properties", "NM.SwAddin.Manufacturing", "NM.SwAddin.Assembly",
    "NM.SwAddin.Geometry", "NM.SwAddin.UI", "NM.SwAddin.Data", "NM.SwAddin.SheetMetal",
    "NM.SwAddin.Import", "NM.SwAddin.Interop", "NM.SwAddin.AssemblyProcessing",
    "SolidWorks.Interop.sldworks", "SolidWorks.Interop.swconst", "SolidWorks.Interop.swpublished"
)

$unknownUsings = @{}
foreach ($file in $csFiles) {
    $content = Get-Content $file -Raw -ErrorAction SilentlyContinue
    if ($null -eq $content) { continue }

    $usingMatches = [regex]::Matches($content, 'using\s+([A-Za-z][\w\.]+);')
    foreach ($match in $usingMatches) {
        $ns = $match.Groups[1].Value
        # Skip using aliases and static usings
        if ($ns -match "=" -or $ns -match "^static\s") { continue }

        # Check if known
        $isKnown = $knownNamespaces | Where-Object { $ns -eq $_ -or $ns.StartsWith("$_.") }
        if (-not $isKnown) {
            $relativePath = $file -replace [regex]::Escape("$projectRoot\"), ""
            if (-not $unknownUsings.ContainsKey($ns)) {
                $unknownUsings[$ns] = @()
            }
            if ($unknownUsings[$ns] -notcontains $relativePath) {
                $unknownUsings[$ns] += $relativePath
            }
        }
    }
}

if ($unknownUsings.Count -gt 0) {
    # Filter to only show truly unknown (not just sub-namespaces)
    $trulyUnknown = $unknownUsings.Keys | Where-Object {
        $ns = $_
        -not ($knownNamespaces | Where-Object { $ns.StartsWith($_) })
    }

    if ($trulyUnknown.Count -gt 0) {
        Write-Host "  Unknown namespaces: $($trulyUnknown.Count)" -ForegroundColor Yellow
        foreach ($ns in $trulyUnknown | Select-Object -First 5) {
            Write-Host "    - $ns" -ForegroundColor Gray
        }
        $warnings += "Unknown namespaces in using directives"
    } else {
        Write-Host "  All using directives reference known namespaces" -ForegroundColor Green
    }
} else {
    Write-Host "  All using directives reference known namespaces" -ForegroundColor Green
}

# === Summary ===
Write-Host ""
Write-Host "=== Summary ===" -ForegroundColor Cyan

if ($issues.Count -eq 0 -and $warnings.Count -eq 0) {
    Write-Host "All checks passed!" -ForegroundColor Green
    exit 0
} else {
    if ($issues.Count -gt 0) {
        Write-Host "Issues (will cause build failure): $($issues.Count)" -ForegroundColor Red
        foreach ($issue in $issues) {
            Write-Host "  - $issue" -ForegroundColor Red
        }
    }
    if ($warnings.Count -gt 0) {
        Write-Host "Warnings (may cause issues): $($warnings.Count)" -ForegroundColor Yellow
        foreach ($warn in $warnings) {
            Write-Host "  - $warn" -ForegroundColor Yellow
        }
    }

    if ($issues.Count -gt 0) {
        exit 1
    }
    exit 0
}
