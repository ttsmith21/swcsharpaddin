<#
.SYNOPSIS
    Syncs work center hourly rates from ProfitKey ERP to nm-config.json.
.DESCRIPTION
    Connects to ProfitKey via ODBC, queries current work center rates,
    and updates the workCenters section of config/nm-config.json.
    Designed to run monthly via Windows Task Scheduler.
.PARAMETER DryRun
    Show what would change without writing to nm-config.json.
.PARAMETER Verbose
    Log every row returned from the ERP query.
.EXAMPLE
    .\sync-erp-rates.ps1 -DryRun
    .\sync-erp-rates.ps1
#>

[CmdletBinding()]
param(
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not $projectRoot) { $projectRoot = (Get-Location).Path }

$syncConfigPath = Join-Path $projectRoot "config\erp-sync.json"
$nmConfigPath   = Join-Path $projectRoot "config\nm-config.json"
$logDir         = Join-Path $projectRoot "logs"
$logPath        = Join-Path $logDir "erp-sync.log"

# --- Logging ---

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message"

    # Always write to console
    switch ($Level) {
        "ERROR"   { Write-Host $entry -ForegroundColor Red }
        "WARNING" { Write-Host $entry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $entry -ForegroundColor Green }
        default   { Write-Host $entry }
    }

    # Append to log file
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $entry | Out-File -FilePath $logPath -Append -Encoding UTF8
}

# --- Main ---

Write-Host "=== ERP Rate Sync ===" -ForegroundColor Cyan
if ($DryRun) { Write-Host "[DRY RUN] No files will be modified." -ForegroundColor Yellow }
Write-Log "Starting ERP rate sync$(if ($DryRun) { ' (dry run)' })"

# 1. Load sync config
if (-not (Test-Path $syncConfigPath)) {
    Write-Log "Sync config not found: $syncConfigPath" "ERROR"
    Write-Log "Create config/erp-sync.json with your ODBC DSN, SQL query, and work center mapping." "ERROR"
    exit 1
}

$syncConfig = Get-Content $syncConfigPath -Raw | ConvertFrom-Json
Write-Log "Loaded sync config from $syncConfigPath"

# Validate sync config has required fields
if (-not $syncConfig.odbc) {
    Write-Log "Missing 'odbc' section in erp-sync.json" "ERROR"; exit 1
}
if (-not $syncConfig.query) {
    Write-Log "Missing 'query' in erp-sync.json" "ERROR"; exit 1
}
if (-not $syncConfig.columnMapping) {
    Write-Log "Missing 'columnMapping' in erp-sync.json" "ERROR"; exit 1
}
if (-not $syncConfig.workCenterMapping) {
    Write-Log "Missing 'workCenterMapping' in erp-sync.json" "ERROR"; exit 1
}

# 2. Build ODBC connection string
$connString = $syncConfig.odbc.connectionString
if (-not $connString) {
    if (-not $syncConfig.odbc.dsn) {
        Write-Log "No DSN or connectionString specified in erp-sync.json" "ERROR"; exit 1
    }
    $connString = "DSN=$($syncConfig.odbc.dsn)"
}
Write-Log "ODBC connection: $connString"

# 3. Connect and query
$idColumn   = $syncConfig.columnMapping.idColumn
$rateColumn = $syncConfig.columnMapping.rateColumn
$query      = $syncConfig.query

Write-Log "Executing query: $query"

$connection = $null
$command    = $null
$reader     = $null

try {
    $connection = New-Object System.Data.Odbc.OdbcConnection($connString)
    $connection.Open()
    Write-Log "Connected to ERP database"

    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $command.CommandTimeout = 30
    $reader = $command.ExecuteReader()

    # 4. Read rows and build rate map
    $erpRates = @{}
    $rowCount = 0

    while ($reader.Read()) {
        $rowCount++
        $wcId = $reader[$idColumn].ToString().Trim()
        $rateRaw = $reader[$rateColumn]

        $rate = 0.0
        if (-not [double]::TryParse($rateRaw.ToString(), [ref]$rate)) {
            Write-Log "Row $rowCount : Cannot parse rate '$rateRaw' for work center '$wcId' - skipping" "WARNING"
            continue
        }

        if ($rate -le 0) {
            Write-Log "Row $rowCount : Rate $rate for '$wcId' is not positive - skipping" "WARNING"
            continue
        }

        $erpRates[$wcId] = $rate
        if ($VerbosePreference -eq 'Continue') {
            Write-Log "Row $rowCount : $wcId = `$$rate/hr"
        }
    }

    Write-Log "Query returned $rowCount rows, $($erpRates.Count) valid rates"

} catch {
    Write-Log "ODBC error: $($_.Exception.Message)" "ERROR"
    exit 1
} finally {
    if ($reader)     { $reader.Close() }
    if ($command)    { $command.Dispose() }
    if ($connection) { $connection.Close(); $connection.Dispose() }
}

if ($erpRates.Count -eq 0) {
    Write-Log "No valid rates returned from query. Nothing to update." "WARNING"
    exit 0
}

# 5. Load nm-config.json
if (-not (Test-Path $nmConfigPath)) {
    Write-Log "nm-config.json not found: $nmConfigPath" "ERROR"; exit 1
}

$nmConfigRaw = Get-Content $nmConfigPath -Raw
$nmConfig = $nmConfigRaw | ConvertFrom-Json
Write-Log "Loaded nm-config.json"

# 6. Map ERP IDs to nm-config keys and apply updates
$wcMapping = @{}
$syncConfig.workCenterMapping.PSObject.Properties | ForEach-Object {
    $wcMapping[$_.Name] = $_.Value
}

$updatedCount  = 0
$skippedCount  = 0
$unchangedCount = 0
$changes = @()

foreach ($erpId in $erpRates.Keys) {
    if (-not $wcMapping.ContainsKey($erpId)) {
        Write-Log "ERP ID '$erpId' has no mapping in workCenterMapping - skipping" "WARNING"
        $skippedCount++
        continue
    }

    $configKey = $wcMapping[$erpId]
    $newRate = $erpRates[$erpId]

    # Check if the key exists in nm-config workCenters
    $currentRate = $null
    $prop = $nmConfig.workCenters.PSObject.Properties | Where-Object { $_.Name -eq $configKey }
    if ($prop) {
        $currentRate = $prop.Value
    }

    if ($null -eq $currentRate) {
        Write-Log "Config key '$configKey' (ERP: $erpId) not found in nm-config.json workCenters - skipping" "WARNING"
        $skippedCount++
        continue
    }

    if ([Math]::Abs($currentRate - $newRate) -lt 0.001) {
        $unchangedCount++
        if ($VerbosePreference -eq 'Continue') {
            Write-Log "$configKey : unchanged at `$$newRate/hr"
        }
        continue
    }

    $changes += [PSCustomObject]@{
        Key     = $configKey
        ErpId   = $erpId
        OldRate = $currentRate
        NewRate = $newRate
    }

    if (-not $DryRun) {
        $nmConfig.workCenters.$configKey = $newRate
    }
    $updatedCount++
}

# 7. Report changes
if ($changes.Count -gt 0) {
    Write-Host "`nRate changes:" -ForegroundColor Cyan
    foreach ($c in $changes) {
        $arrow = if ($c.NewRate -gt $c.OldRate) { "^" } else { "v" }
        $color = if ($c.NewRate -gt $c.OldRate) { "Yellow" } else { "Green" }
        $msg = "  $($c.Key) : `$$($c.OldRate) -> `$$($c.NewRate) $arrow"
        Write-Host $msg -ForegroundColor $color
        Write-Log "$($c.Key) ($($c.ErpId)): $($c.OldRate) -> $($c.NewRate)"
    }
} else {
    Write-Host "`nNo rate changes detected." -ForegroundColor Green
}

Write-Log "Summary: $updatedCount updated, $unchangedCount unchanged, $skippedCount skipped"

# 8. Write updated config
if (-not $DryRun -and $updatedCount -gt 0) {
    # Backup
    $backupPath = "$nmConfigPath.bak"
    Copy-Item -Path $nmConfigPath -Destination $backupPath -Force
    Write-Log "Backed up nm-config.json to nm-config.json.bak"

    # Add/update lastErpSync timestamp
    $syncTimestamp = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $existingProp = $nmConfig.PSObject.Properties | Where-Object { $_.Name -eq "lastErpSync" }
    if ($existingProp) {
        $nmConfig.lastErpSync = $syncTimestamp
    } else {
        $nmConfig | Add-Member -NotePropertyName "lastErpSync" -NotePropertyValue $syncTimestamp
    }

    # Write with formatting preserved (ConvertTo-Json with sufficient depth)
    $nmConfig | ConvertTo-Json -Depth 10 | Set-Content -Path $nmConfigPath -Encoding UTF8
    Write-Log "Updated nm-config.json with $updatedCount new rates (lastErpSync: $syncTimestamp)" "SUCCESS"
} elseif ($DryRun -and $updatedCount -gt 0) {
    Write-Log "DRY RUN complete. $updatedCount rate(s) would be updated." "SUCCESS"
} else {
    Write-Log "No updates needed." "SUCCESS"
}

Write-Host "`n=== Sync Complete ===" -ForegroundColor Cyan
exit 0
