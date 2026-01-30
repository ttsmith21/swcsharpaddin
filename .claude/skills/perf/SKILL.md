---
name: perf
description: Analyze timing data from QA runs to identify performance bottlenecks and track regressions. Run after /qa to see what's slow.
user_invocable: true
---

# Performance Analysis Skill

Analyzes timing data from QA runs to identify bottlenecks and track regressions.

## What This Does

1. Reads the latest timing.csv from tests/Run_Latest/ or most recent timestamped run
2. Summarizes timing by operation category
3. Identifies the slowest operations
4. Compares against baseline (if exists)
5. Reports regressions and improvements

## Commands to Execute

```powershell
# Step 1: Find latest run folder
$testsDir = "C:\Users\tsmith\source\repos\swcsharpaddin\tests"
$latestRun = Get-ChildItem $testsDir -Directory |
    Where-Object { $_.Name -match "Run_\d{8}_\d{6}" } |
    Sort-Object Name -Descending |
    Select-Object -First 1

if (-not $latestRun) {
    Write-Host "No QA run found. Run /qa first." -ForegroundColor Red
    return
}

Write-Host "Analyzing: $($latestRun.FullName)" -ForegroundColor Cyan

# Step 2: Read timing CSV
$timingCsv = Join-Path $latestRun.FullName "timing.csv"
if (Test-Path $timingCsv) {
    $timings = Import-Csv $timingCsv

    Write-Host ""
    Write-Host "=== Top 10 Slowest Operations (by total time) ===" -ForegroundColor Yellow
    $grouped = $timings | Group-Object Name | ForEach-Object {
        [PSCustomObject]@{
            Name = $_.Name
            Count = $_.Count
            TotalMs = [math]::Round(($_.Group | Measure-Object -Property ElapsedMs -Sum).Sum, 1)
            AvgMs = [math]::Round(($_.Group | Measure-Object -Property ElapsedMs -Average).Average, 1)
        }
    } | Sort-Object TotalMs -Descending | Select-Object -First 10

    $grouped | Format-Table -AutoSize

    Write-Host ""
    Write-Host "=== By Category ===" -ForegroundColor Yellow

    # InsertBends operations
    $insertBends = $timings | Where-Object { $_.Name -like "InsertBends2_*" }
    if ($insertBends) {
        $total = [math]::Round(($insertBends | Measure-Object -Property ElapsedMs -Sum).Sum, 0)
        Write-Host "InsertBends2_*    : ${total}ms total ($($insertBends.Count) calls)"
    }

    # Flatten operations
    $flatten = $timings | Where-Object { $_.Name -like "TryFlatten_*" }
    if ($flatten) {
        $total = [math]::Round(($flatten | Measure-Object -Property ElapsedMs -Sum).Sum, 0)
        Write-Host "TryFlatten_*      : ${total}ms total ($($flatten.Count) calls)"
    }

    # Geometry operations
    $geometry = $timings | Where-Object { $_.Name -match "GetLargestFace|FindLongestLinearEdge" }
    if ($geometry) {
        $total = [math]::Round(($geometry | Measure-Object -Property ElapsedMs -Sum).Sum, 0)
        Write-Host "Geometry          : ${total}ms total ($($geometry.Count) calls)"
    }

    # Property operations
    $props = $timings | Where-Object { $_.Name -like "CustomProperty_*" }
    if ($props) {
        $total = [math]::Round(($props | Measure-Object -Property ElapsedMs -Sum).Sum, 0)
        Write-Host "CustomProperty_*  : ${total}ms total ($($props.Count) calls)"
    }

    # Batch scope
    $batch = $timings | Where-Object { $_.Name -eq "BatchPerformanceScope" }
    if ($batch) {
        $total = [math]::Round(($batch | Measure-Object -Property ElapsedMs -Sum).Sum, 0)
        Write-Host "BatchPerfScope    : ${total}ms total ($($batch.Count) calls)"
    }

} else {
    Write-Host "No timing.csv found in $($latestRun.FullName)" -ForegroundColor Red
    Write-Host "Ensure PerformanceTracker is enabled (EnablePerformanceMonitoring = true)"
}

# Step 3: Check for baseline
$baselinePath = Join-Path $testsDir "timing-baseline.json"
Write-Host ""
if (Test-Path $baselinePath) {
    Write-Host "=== Baseline Status ===" -ForegroundColor Yellow
    Write-Host "Baseline exists at: $baselinePath"
    Write-Host "Review results.json for regression flags"
} else {
    Write-Host "=== No Baseline Found ===" -ForegroundColor Yellow
    Write-Host "To create a baseline after a good run:"
    Write-Host "1. Run /qa to generate timing.csv"
    Write-Host "2. Create timing-baseline.json with timer averages"
}
```

## Interpreting Results

**Key timers to watch:**
- `InsertBends2_*` - Sheet metal conversion (target: <500ms each)
- `TryFlatten_*` - Flattening operations (target: <200ms each)
- `CustomProperty_*` - Property I/O (target: <50ms each)
- `BatchPerformanceScope` - Graphics disable overhead (target: <100ms)

**Red flags:**
- Any operation averaging >1000ms
- InsertBends taking >2000ms (geometry complexity issue)
- Property operations taking >100ms (network path issue)

## Creating a Baseline

After a successful QA run with acceptable performance:

```powershell
# Create baseline JSON manually from timing data
# Format:
# {
#   "CreatedAt": "2026-01-30T12:00:00Z",
#   "RunId": "20260130_120000",
#   "FileCount": 16,
#   "TotalElapsedMs": 30000,
#   "Timers": {
#     "InsertBends2_Probe": { "AvgMs": 150.0, "MaxMs": 300.0, "Count": 10 },
#     "TryFlatten_Final": { "AvgMs": 80.0, "MaxMs": 150.0, "Count": 10 }
#   }
# }
```

## When to Use

- After `/qa` to understand where time is spent
- When investigating slow processing
- Before and after optimization attempts
- To establish performance baseline for regression tracking
