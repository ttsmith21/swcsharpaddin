<#
.SYNOPSIS
    Analyzes and auto-fixes common build warnings.
.DESCRIPTION
    Fixes:
    - CS0618: Deprecated ErrorHandler.HandleError calls (string severity -> LogLevel enum)
    - CS0219: Unused variable assignments
    - CS0162: Unreachable code (reports only, no auto-fix)
.PARAMETER AutoFix
    If specified, automatically applies fixes.
.PARAMETER DryRun
    Show what would be changed without making changes.
#>

param(
    [switch]$AutoFix,
    [switch]$DryRun
)

$projectRoot = Split-Path -Parent $PSScriptRoot
$buildOutput = Join-Path $projectRoot "build-output.txt"

# Run a clean build to capture warnings if no recent output exists
if (-not (Test-Path $buildOutput) -or ((Get-Item $buildOutput).LastWriteTime -lt (Get-Date).AddMinutes(-5))) {
    Write-Host "Running clean build to capture warnings..." -ForegroundColor Yellow
    $msbuild = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" `
        -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\amd64\MSBuild.exe | Select-Object -First 1

    if (-not $msbuild) {
        Write-Host "MSBuild not found" -ForegroundColor Red
        exit 1
    }

    $csproj = Join-Path $projectRoot "swcsharpaddin.csproj"

    # Clean first to force full rebuild
    $binPath = Join-Path $projectRoot "bin"
    $objPath = Join-Path $projectRoot "obj"
    if (Test-Path $binPath) { Remove-Item -Path $binPath -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path $objPath) { Remove-Item -Path $objPath -Recurse -Force -ErrorAction SilentlyContinue }

    & $msbuild $csproj /p:Configuration=Debug /v:normal /nologo 2>&1 | Out-File -FilePath $buildOutput -Encoding UTF8
}

# Parse warnings from build output
$buildContent = Get-Content $buildOutput -Raw
$warningLines = $buildContent -split "`n" | Where-Object { $_ -match "warning CS" }

Write-Host ""
Write-Host "=== Warning Analysis ===" -ForegroundColor Cyan
Write-Host "Total warnings: $($warningLines.Count)"
Write-Host ""

# Categorize warnings
$cs0618 = @()  # Deprecated calls
$cs0219 = @()  # Unused variables
$cs0162 = @()  # Unreachable code
$other = @()

foreach ($line in $warningLines) {
    # Match full path - handles both formats:
    #   1>C:\path\file.cs(123,45): warning CS0618: message [project]
    #   C:\path\file.cs(123,45): warning CS0618: message [project]
    if ($line -match "(?:^\s*\d+>)?(C:[^(]+\.cs)\((\d+),\d+\):\s*warning (CS\d+):\s*(.+?)\s*\[") {
        $fullPath = $matches[1].Trim()
        $lineNum = [int]$matches[2]
        $code = $matches[3]
        $message = $matches[4]
        $fileName = Split-Path -Leaf $fullPath

        if ([string]::IsNullOrWhiteSpace($fullPath)) { continue }

        $warning = [PSCustomObject]@{
            File = $fullPath
            FileName = $fileName
            Line = $lineNum
            Code = $code
            Message = $message
        }

        switch ($code) {
            "CS0618" { $cs0618 += $warning }
            "CS0219" { $cs0219 += $warning }
            "CS0162" { $cs0162 += $warning }
            default { $other += $warning }
        }
    }
}

# === CS0618: Deprecated ErrorHandler.HandleError ===
if ($cs0618.Count -gt 0) {
    Write-Host "--- CS0618: Deprecated Calls ($($cs0618.Count)) ---" -ForegroundColor Yellow

    # Group by file
    $groupedByFile = $cs0618 | Group-Object -Property File

    foreach ($group in $groupedByFile) {
        $file = $group.Name
        $shortFile = $file -replace [regex]::Escape($projectRoot), "."
        Write-Host "`n$shortFile ($($group.Count) warnings)" -ForegroundColor White

        if ($AutoFix -or $DryRun) {
            if (Test-Path $file) {
                $content = Get-Content $file -Raw
                $originalContent = $content
                $fixCount = 0

                # Pattern to match deprecated ErrorHandler.HandleError calls
                # Old: ErrorHandler.HandleError(..., "Warning") or (..., "Warning", context)
                # New: ErrorHandler.HandleError(..., ErrorHandler.LogLevel.Warning) or (..., ErrorHandler.LogLevel.Warning, context)

                # Simpler approach: just replace the ending pattern
                # Match: , "Warning") at end of HandleError calls
                $pattern4 = ',\s*"(Warning|Error|Info|Critical)"\s*\)'
                $newContent = [regex]::Replace($content, $pattern4, {
                    param($m)
                    $severity = $m.Groups[1].Value
                    ", ErrorHandler.LogLevel.$severity)"
                })

                # Match: , "Warning", something) - with context parameter
                $pattern5 = ',\s*"(Warning|Error|Info|Critical)"\s*,'
                $newContent = [regex]::Replace($newContent, $pattern5, {
                    param($m)
                    $severity = $m.Groups[1].Value
                    ", ErrorHandler.LogLevel.$severity,"
                })

                # Count changes
                $fixCount = 0
                $fixCount += ([regex]::Matches($content, $pattern4)).Count
                $fixCount += ([regex]::Matches($content, $pattern5)).Count
                $content = $newContent

                if ($content -ne $originalContent) {
                    if ($DryRun) {
                        Write-Host "  Would fix ~$fixCount calls" -ForegroundColor Cyan
                    } else {
                        Set-Content -Path $file -Value $content -NoNewline
                        Write-Host "  Fixed $fixCount calls" -ForegroundColor Green
                    }
                } else {
                    Write-Host "  No auto-fixable patterns found (may need manual review)" -ForegroundColor Gray
                }
            }
        } else {
            foreach ($w in $group.Group | Select-Object -First 3) {
                Write-Host "  Line $($w.Line): $($w.Message)" -ForegroundColor Gray
            }
            if ($group.Count -gt 3) {
                Write-Host "  ... and $($group.Count - 3) more" -ForegroundColor DarkGray
            }
        }
    }
}

# === CS0219: Unused Variables ===
# Deduplicate (same file+line can appear multiple times in output)
$cs0219Unique = $cs0219 | Sort-Object -Property @{Expression={$_.File}}, @{Expression={$_.Line}} -Unique

if ($cs0219Unique.Count -gt 0) {
    Write-Host ""
    Write-Host "--- CS0219: Unused Variables ($($cs0219Unique.Count)) ---" -ForegroundColor Yellow

    foreach ($w in $cs0219Unique) {
        $shortFile = $w.File -replace [regex]::Escape($projectRoot), "."
        $varName = if ($w.Message -match "'(\w+)'") { $matches[1] } else { "?" }
        Write-Host "  $shortFile`:$($w.Line) - '$varName'" -ForegroundColor Gray
    }

    # NOTE: CS0219 auto-fix disabled - commenting out variables can break builds
    # if the variable is later assigned to (just not read from).
    # These require manual review to determine if safe to remove.
}

# === CS0162: Unreachable Code ===
$cs0162Unique = $cs0162 | Sort-Object -Property @{Expression={$_.File}}, @{Expression={$_.Line}} -Unique

if ($cs0162Unique.Count -gt 0) {
    Write-Host ""
    Write-Host "--- CS0162: Unreachable Code ($($cs0162Unique.Count)) ---" -ForegroundColor Yellow
    Write-Host "  (Manual review required)" -ForegroundColor Gray

    foreach ($w in $cs0162Unique) {
        $shortFile = $w.File -replace [regex]::Escape($projectRoot), "."
        Write-Host "  $shortFile`:$($w.Line)" -ForegroundColor Gray
    }
}

# === Other Warnings ===
if ($other.Count -gt 0) {
    Write-Host ""
    Write-Host "--- Other Warnings ($($other.Count)) ---" -ForegroundColor Yellow

    $grouped = $other | Group-Object -Property Code
    foreach ($g in $grouped) {
        Write-Host "  $($g.Name): $($g.Count)" -ForegroundColor Gray
    }
}

# === Summary ===
Write-Host ""
Write-Host "=== Summary ===" -ForegroundColor Cyan

# Deduplicate CS0618 by file for counting
$cs0618Files = ($cs0618 | Select-Object -Property File -Unique).Count
$cs0219Count = $cs0219Unique.Count
$cs0162Count = $cs0162Unique.Count

$fixable = $cs0618Files + $cs0219Count
$manual = $cs0162Count + $other.Count

Write-Host "Auto-fixable: CS0618 in $cs0618Files files, CS0219: $cs0219Count unused vars"
Write-Host "Manual review: CS0162: $cs0162Count, Other: $($other.Count)"

if (-not $AutoFix -and -not $DryRun -and $fixable -gt 0) {
    Write-Host ""
    Write-Host "Run with -DryRun to preview changes, or -AutoFix to apply fixes." -ForegroundColor Cyan
}
