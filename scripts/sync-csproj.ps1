<#
.SYNOPSIS
    Syncs .cs files on disk with the csproj file (old-style .NET Framework projects).

.DESCRIPTION
    Scans for .cs files in the project directory and adds any missing ones to the csproj.
    This prevents the common issue where files exist on disk but aren't compiled.

.PARAMETER DryRun
    If specified, only reports what would be added without modifying the csproj.
#>
param(
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$csprojPath = Join-Path $projectRoot "swcsharpaddin.csproj"

Write-Host "=== sync-csproj ===" -ForegroundColor Cyan
Write-Host "Project root: $projectRoot"
Write-Host "Csproj: $csprojPath"
if ($DryRun) { Write-Host "[DRY RUN MODE]" -ForegroundColor Yellow }
Write-Host ""

# Directories to exclude (can appear at any level in the path)
$excludeDirs = @("obj", "bin", "Backup", "NM.Core.Tests", "NM.StepClassifierAddin", ".git", ".claude")

# Find all .cs files on disk
$allCsFiles = Get-ChildItem -Path $projectRoot -Filter "*.cs" -Recurse |
    Where-Object {
        $relativePath = $_.FullName.Substring($projectRoot.Length + 1)
        $pathParts = $relativePath -split '\\|/'
        # Check if any part of the path is in the exclude list
        $excluded = $false
        foreach ($part in $pathParts) {
            if ($excludeDirs -contains $part) {
                $excluded = $true
                break
            }
        }
        -not $excluded
    } |
    ForEach-Object {
        $_.FullName.Substring($projectRoot.Length + 1).Replace("/", "\")
    }

Write-Host "Found $($allCsFiles.Count) .cs files on disk (excluding $($excludeDirs -join ', '))"

# Parse csproj to find existing Compile includes
[xml]$csproj = Get-Content $csprojPath
$nsMgr = New-Object System.Xml.XmlNamespaceManager($csproj.NameTable)
$nsMgr.AddNamespace("ms", "http://schemas.microsoft.com/developer/msbuild/2003")

$existingIncludes = @()
$compileNodes = $csproj.SelectNodes("//ms:Compile", $nsMgr)
foreach ($node in $compileNodes) {
    $include = $node.GetAttribute("Include")
    if ($include) {
        $existingIncludes += $include.Replace("/", "\")
    }
}

Write-Host "Found $($existingIncludes.Count) Compile includes in csproj"
Write-Host ""

# Find missing files
$missing = $allCsFiles | Where-Object { $existingIncludes -notcontains $_ }

if ($missing.Count -eq 0) {
    Write-Host "All .cs files are already in csproj!" -ForegroundColor Green
    exit 0
}

Write-Host "Missing from csproj ($($missing.Count) files):" -ForegroundColor Yellow
foreach ($file in $missing) {
    Write-Host "  + $file"
}
Write-Host ""

if ($DryRun) {
    Write-Host "[DRY RUN] Would add $($missing.Count) files to csproj" -ForegroundColor Yellow
    exit 0
}

# Find the ItemGroup containing Compile elements
$compileItemGroup = $csproj.SelectSingleNode("//ms:ItemGroup[ms:Compile]", $nsMgr)
if (-not $compileItemGroup) {
    Write-Host "ERROR: Could not find ItemGroup with Compile elements" -ForegroundColor Red
    exit 1
}

# Add missing files
foreach ($file in $missing) {
    $newNode = $csproj.CreateElement("Compile", "http://schemas.microsoft.com/developer/msbuild/2003")
    $newNode.SetAttribute("Include", $file)

    # Special handling for Forms
    if ($file -match "Form\.cs$" -or $file -match "Dialog\.cs$") {
        $subType = $csproj.CreateElement("SubType", "http://schemas.microsoft.com/developer/msbuild/2003")
        $subType.InnerText = "Form"
        $newNode.AppendChild($subType) | Out-Null
    }

    $compileItemGroup.AppendChild($newNode) | Out-Null
}

# Save the csproj
$csproj.Save($csprojPath)

Write-Host "Added $($missing.Count) files to csproj" -ForegroundColor Green
Write-Host ""
Write-Host "=== sync-csproj complete ===" -ForegroundColor Cyan
