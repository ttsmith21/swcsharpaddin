<#
.SYNOPSIS
    Parses Import.prn (VBA gold standard) and generates manifest-v2.json with enriched baseline data.
.DESCRIPTION
    Reads the VBA-produced Import.prn file, extracts per-part data (descriptions, stock IDs,
    flat areas, routing operations, BOM quantities), maps part numbers to test file names,
    and merges with existing manifest.json to produce manifest-v2.json.
.PARAMETER PrnPath
    Path to Import.prn file. Default: tests/GoldStandard_Baseline/Import_GOLD_STANDARD_ASM_CLEAN.prn
.PARAMETER ManifestPath
    Path to existing manifest.json. Default: tests/GoldStandard_Baseline/manifest.json
.PARAMETER InputDir
    Path to GoldStandard_Inputs directory (for filename mapping). Default: tests/GoldStandard_Inputs
.PARAMETER OutputPath
    Path to write manifest-v2.json. Default: tests/GoldStandard_Baseline/manifest-v2.json
#>
param(
    [string]$PrnPath = "$PSScriptRoot\..\tests\GoldStandard_Baseline\Import_GOLD_STANDARD_ASM_CLEAN.prn",
    [string]$ManifestPath = "$PSScriptRoot\..\tests\GoldStandard_Baseline\manifest.json",
    [string]$InputDir = "$PSScriptRoot\..\tests\GoldStandard_Inputs",
    [string]$OutputPath = "$PSScriptRoot\..\tests\GoldStandard_Baseline\manifest-v2.json"
)

$ErrorActionPreference = "Stop"

# --- Parse Import.prn ---
Write-Host "Parsing Import.prn..." -ForegroundColor Cyan
if (-not (Test-Path $PrnPath)) {
    Write-Error "Import.prn not found at: $PrnPath"
    exit 1
}

$lines = Get-Content $PrnPath
$records = @{}  # keyed by part number
$currentSection = ""
$fieldNames = @()
$inData = $false

function Parse-QuotedLine {
    param([string]$line)
    $tokens = @()
    $current = ""
    $inQuotes = $false
    $hadQuotes = $false  # Track if this token was quoted (to preserve empty strings)
    for ($i = 0; $i -lt $line.Length; $i++) {
        $ch = $line[$i]
        if ($ch -eq '\' -and ($i + 1) -lt $line.Length -and $line[$i + 1] -eq '"') {
            $current += '"'
            $i++
            continue
        }
        if ($ch -eq '"') {
            $inQuotes = -not $inQuotes
            $hadQuotes = $true
            continue
        }
        if (-not $inQuotes -and ($ch -eq ' ' -or $ch -eq "`t")) {
            if ($current.Length -gt 0 -or $hadQuotes) {
                $tokens += $current
                $current = ""
                $hadQuotes = $false
            }
            continue
        }
        $current += $ch
    }
    if ($current.Length -gt 0 -or $hadQuotes) { $tokens += $current }
    return $tokens
}

function Get-OrCreate {
    param([hashtable]$ht, [string]$key)
    if (-not $ht.ContainsKey($key)) {
        $ht[$key] = @{
            PartNumber = $key; Drawing = ""; Description = ""; Revision = ""
            ImType = 0; ImClass = 0; Commodity = ""; StdLot = 0
            OptiMaterial = ""; RawWeight = 0.0; F300Length = 0.0
            BomQuantity = 0; ParentPartNumber = ""; PieceNumber = ""
            Routing = @(); RoutingNotes = @{}
        }
    }
    return $ht[$key]
}

foreach ($rawLine in $lines) {
    $line = $rawLine.Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { $inData = $false; continue }

    if ($line.StartsWith("DECL(")) {
        $currentSection = $line.Substring(5, 2)
        $fieldNames = @()
        $inData = $false
        # Extract field names from DECL line (after DECL(XX) and ADD)
        $tokens = Parse-QuotedLine $line
        foreach ($tok in $tokens) {
            if ($tok -match "^DECL\(" -or $tok -eq "ADD") { continue }
            $fieldNames += $tok
        }
        continue
    }
    if ($line -eq "END") { $inData = $true; continue }
    if (-not $inData) { continue }

    # Parse data row
    $values = Parse-QuotedLine $line
    $data = @{}
    for ($j = 0; $j -lt [Math]::Min($fieldNames.Count, $values.Count); $j++) {
        $data[$fieldNames[$j]] = $values[$j]
    }

    switch ($currentSection) {
        "IM" {
            $key = $data["IM-KEY"]
            if (-not $key) { continue }
            $rec = Get-OrCreate $records $key
            $rec.Drawing = if ($data["IM-DRAWING"]) { $data["IM-DRAWING"] } else { "" }
            $rec.Description = if ($data["IM-DESCR"]) { $data["IM-DESCR"] } else { "" }
            $rec.Revision = if ($data["IM-REV"]) { $data["IM-REV"] } else { "" }
            $rec.ImType = if ($data["IM-TYPE"]) { [int]$data["IM-TYPE"] } else { 0 }
            $rec.ImClass = if ($data["IM-CLASS"]) { [int]$data["IM-CLASS"] } else { 0 }
            $rec.Commodity = if ($data["IM-COMMODITY"]) { $data["IM-COMMODITY"] } else { "" }
            $rec.StdLot = if ($data["IM-STD-LOT"]) { [int]$data["IM-STD-LOT"] } else { 0 }
        }
        "PS" {
            if ($data.ContainsKey("PS-DIM-1")) {
                # Material relationship section
                $key = $data["PS-PARENT-KEY"]
                if (-not $key) { continue }
                $rec = Get-OrCreate $records $key
                $rec.OptiMaterial = if ($data["PS-SUBORD-KEY"]) { $data["PS-SUBORD-KEY"] } else { "" }
                $rec.RawWeight = if ($data["PS-QTY-P"]) { [double]$data["PS-QTY-P"] } else { 0.0 }
                $rec.F300Length = if ($data["PS-DIM-1"]) { [double]$data["PS-DIM-1"] } else { 0.0 }
            } else {
                # BOM section
                $key = $data["PS-SUBORD-KEY"]
                if (-not $key) { continue }
                $rec = Get-OrCreate $records $key
                $rec.ParentPartNumber = if ($data["PS-PARENT-KEY"]) { $data["PS-PARENT-KEY"] } else { "" }
                $rec.PieceNumber = if ($data["PS-PIECE-NO"]) { $data["PS-PIECE-NO"] } else { "" }
                $rec.BomQuantity = if ($data["PS-QTY-P"]) { [double]$data["PS-QTY-P"] } else { 0 }
            }
        }
        "RT" {
            $key = $data["RT-ITEM-KEY"]
            if (-not $key) { continue }
            $rec = Get-OrCreate $records $key
            $rec.Routing += @{
                WorkCenter = if ($data["RT-WORKCENTER-KEY"]) { $data["RT-WORKCENTER-KEY"].Trim() } else { "" }
                OpNumber = if ($data["RT-OP-NUM"]) { [int]$data["RT-OP-NUM"] } else { 0 }
                Setup = if ($data["RT-SETUP"]) { [double]$data["RT-SETUP"] } else { 0.0 }
                Run = if ($data["RT-RUN-STD"]) { [double]$data["RT-RUN-STD"] } else { 0.0 }
            }
        }
        "RN" {
            $key = $data["RN-ITEM-KEY"]
            if (-not $key) { continue }
            $rec = Get-OrCreate $records $key
            $opNum = if ($data["RN-OP-NUM"]) { [int]$data["RN-OP-NUM"] } else { 0 }
            $noteText = if ($data["RN-DESCR"]) { $data["RN-DESCR"] } else { "" }
            if (-not $rec.RoutingNotes.ContainsKey($opNum)) {
                $rec.RoutingNotes[$opNum] = @()
            }
            $rec.RoutingNotes[$opNum] += $noteText
        }
    }
}

Write-Host "  Parsed $($records.Count) parts from Import.prn" -ForegroundColor Green

# --- Map PRN part numbers to test file names ---
Write-Host "Mapping part numbers to test files..." -ForegroundColor Cyan
$inputFiles = Get-ChildItem $InputDir -Recurse -Include "*.sldprt","*.SLDPRT","*.sldasm","*.SLDASM" -File
$fileMap = @{}  # PRN part number -> test file name (with extension)

foreach ($prnKey in $records.Keys) {
    # Skip the assembly parent record
    if ($prnKey -eq "GOLD_STANDARD_ASM_CLEAN") { continue }
    # Skip CUST- prefixed records (duplicates of D11, D3)
    if ($prnKey.StartsWith("CUST-")) { continue }

    # Try exact match first (without extension)
    $match = $inputFiles | Where-Object { $_.BaseName -eq $prnKey } | Select-Object -First 1
    if (-not $match) {
        # Try prefix match (PRN key is the prefix of the filename)
        $match = $inputFiles | Where-Object { $_.BaseName.StartsWith($prnKey) } | Select-Object -First 1
    }
    if ($match) {
        $fileMap[$prnKey] = $match.Name
    } else {
        Write-Host "  WARNING: No input file found for PRN key '$prnKey'" -ForegroundColor Yellow
    }
}
Write-Host "  Mapped $($fileMap.Count) parts to test files" -ForegroundColor Green

# --- Load existing manifest.json ---
Write-Host "Loading existing manifest.json..." -ForegroundColor Cyan
$existingManifest = Get-Content $ManifestPath -Raw | ConvertFrom-Json

# --- Build manifest-v2.json ---
Write-Host "Building manifest-v2.json..." -ForegroundColor Cyan

# Known deviations between VBA and C# (documented)
$tubeDeviations = @{
    "C2_RectTube_2x1.SLDPRT" = "TubeGeometryExtractor handles rectangular tubes. VBA classified as Generic."
    "C3_SquareTube_2x2.SLDPRT" = "TubeGeometryExtractor handles square tubes. VBA classified as Generic."
    "C4_AngleIron_2x2.SLDPRT" = "TubeGeometryExtractor handles angle profiles. VBA classified as Generic."
    "C5_CChannel.SLDPRT" = "TubeGeometryExtractor handles C-channel. VBA classified as Generic."
    "C6_IBeam.SLDPRT" = "TubeGeometryExtractor handles I-beam. VBA classified as Generic."
}

$manifestV2 = [ordered]@{
    version = "2.0"
    description = "Gold Standard expected outcomes with VBA baseline data from Import.prn"
    generatedAt = (Get-Date -Format "o")
    defaultTolerances = [ordered]@{
        thickness = 0.001
        mass = 0.01
        dimensions = 0.001
        cost = 0.05
        routing = 0.01
        flatArea = 0.02
    }
    files = [ordered]@{}
}

# Process each file from existing manifest
foreach ($prop in $existingManifest.files.PSObject.Properties) {
    $fileName = $prop.Name
    if ($fileName.StartsWith("_comment")) { continue }

    $existing = $prop.Value
    $entry = [ordered]@{}

    # Copy existing top-level fields
    $entry.shouldPass = $existing.shouldPass
    if ($existing.expectedFailureReason) { $entry.expectedFailureReason = $existing.expectedFailureReason }
    if ($existing.expectedClassification) { $entry.expectedClassification = $existing.expectedClassification }
    if ($null -ne $existing.expectedThickness_in) { $entry.expectedThickness_in = $existing.expectedThickness_in }
    if ($null -ne $existing.expectedBendCount) { $entry.expectedBendCount = $existing.expectedBendCount }
    if ($null -ne $existing.expectedTubeOD_in) { $entry.expectedTubeOD_in = $existing.expectedTubeOD_in }
    if ($null -ne $existing.expectedTubeWall_in) { $entry.expectedTubeWall_in = $existing.expectedTubeWall_in }
    if ($null -ne $existing.expectedTubeID_in) { $entry.expectedTubeID_in = $existing.expectedTubeID_in }
    if ($null -ne $existing.expectedTubeLength_in) { $entry.expectedTubeLength_in = $existing.expectedTubeLength_in }

    # Find matching PRN record
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    $prnKey = $null
    foreach ($k in $fileMap.Keys) {
        if ($fileMap[$k] -ieq $fileName) { $prnKey = $k; break }
    }
    # Also try direct match by basename prefix
    if (-not $prnKey) {
        foreach ($k in $records.Keys) {
            if ($baseName.StartsWith($k) -or $k -eq $baseName) { $prnKey = $k; break }
        }
    }

    # Build vbaBaseline from Import.prn data
    if ($prnKey -and $records.ContainsKey($prnKey)) {
        $prn = $records[$prnKey]
        $vba = [ordered]@{}

        if ($prn.Description) { $vba.description = $prn.Description }
        if ($prn.OptiMaterial) { $vba.optiMaterial = $prn.OptiMaterial }
        if ($prn.RawWeight -gt 0) { $vba.rawWeight = [Math]::Round($prn.RawWeight, 4) }
        if ($prn.F300Length -gt 0) { $vba.f300Length = $prn.F300Length }
        if ($prn.BomQuantity -gt 0) { $vba.bomQty = [int]$prn.BomQuantity }

        # Routing operations
        if ($prn.Routing.Count -gt 0) {
            $routing = [ordered]@{}
            foreach ($step in $prn.Routing) {
                $routing[$step.WorkCenter] = [ordered]@{
                    op = $step.OpNumber
                    setup = $step.Setup
                    run = $step.Run
                }
            }
            $vba.routing = $routing
        }

        $entry.vbaBaseline = $vba
    }

    # csharpExpected (empty for now - will be filled as C# reaches parity)
    $entry.csharpExpected = [ordered]@{}

    # Known deviations
    $deviations = [ordered]@{}
    if ($tubeDeviations.ContainsKey($fileName)) {
        $deviations.classification = [ordered]@{
            reason = $tubeDeviations[$fileName]
            status = "INTENTIONAL"
        }
    }
    # All parts: optiMaterial and description not yet generated by C#
    if ($prnKey -and $records.ContainsKey($prnKey)) {
        $prn = $records[$prnKey]
        if ($prn.OptiMaterial -and $existing.shouldPass -eq $true) {
            $deviations.optiMaterial = [ordered]@{
                reason = "OptiMaterial resolution not yet wired in C# pipeline"
                status = "NOT_IMPLEMENTED"
            }
        }
        if ($prn.Description -and $existing.shouldPass -eq $true) {
            $deviations.description = [ordered]@{
                reason = "Description generation not yet implemented in C#"
                status = "NOT_IMPLEMENTED"
            }
        }
        if ($prn.Routing.Count -gt 0 -and $existing.shouldPass -eq $true) {
            $deviations.routing = [ordered]@{
                reason = "Routing not yet generated in QA results"
                status = "NOT_IMPLEMENTED"
            }
        }
    }
    if ($deviations.Count -gt 0) {
        $entry.knownDeviations = $deviations
    }

    # Notes from existing manifest
    if ($existing._note) { $entry._note = $existing._note }
    if ($existing._status) { $entry._status = $existing._status }

    $manifestV2.files[$fileName] = $entry
}

# --- Write manifest-v2.json ---
Write-Host "Writing manifest-v2.json..." -ForegroundColor Cyan
$json = $manifestV2 | ConvertTo-Json -Depth 10
$json | Set-Content $OutputPath -Encoding UTF8

$fileCount = ($manifestV2.files.Keys | Where-Object { -not $_.StartsWith("_") }).Count
$vbaCount = ($manifestV2.files.Values | Where-Object { $_.vbaBaseline }).Count
Write-Host ""
Write-Host "=== manifest-v2.json Generated ===" -ForegroundColor Green
Write-Host "  Total entries: $fileCount"
Write-Host "  With VBA baseline data: $vbaCount"
Write-Host "  Output: $OutputPath"
Write-Host ""

# Print summary of VBA baseline data
Write-Host "VBA Baseline Summary:" -ForegroundColor Cyan
foreach ($fn in $manifestV2.files.Keys) {
    if ($fn.StartsWith("_")) { continue }
    $e = $manifestV2.files[$fn]
    $vbaData = if ($e.vbaBaseline) { "VBA: $($e.vbaBaseline.optiMaterial)" } else { "(no VBA data)" }
    $status = if ($e.shouldPass) { "PASS" } else { "FAIL" }
    $class = if ($e.expectedClassification) { $e.expectedClassification } else { "-" }
    Write-Host ("  {0,-45} {1,-5} {2,-12} {3}" -f $fn, $status, $class, $vbaData)
}
