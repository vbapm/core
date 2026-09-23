# Assess-ExcelInstances.ps1
#
# End-of-suite assessment for the e2e suite. Reconciles the coordination
# registry with the live EXCEL.EXE process list, then reports instances that
# linger longer than expected:
#
#   - "zombie": an entry in the registry's `inactive` list whose PID is STILL
#     alive — it was deactivated (supposedly closed) but the EXCEL.EXE process
#     lingers.
#   - "lingering": a live, automation-owned instance that has been alive longer
#     than -AgeSeconds (i.e. it should have been closed by its run but wasn't).
#
# Each flagged instance is reported together with the workbooks it held (from
# the registry), so the open workbook paths identify which e2e test is
# responsible for the zombie.
#
# The summary + flagged instances are also appended to the shared instance log
# (%TEMP%\Excel-Instances\instances.log or $env:VBA_INSTANCE_LOG).
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\scripts\ps\Assess-ExcelInstances.ps1
#   powershell -ExecutionPolicy Bypass -File .\scripts\ps\Assess-ExcelInstances.ps1 -AgeSeconds 120
#   powershell -ExecutionPolicy Bypass -File .\scripts\ps\Assess-ExcelInstances.ps1 -FailOnZombie
#
# Exit codes:
#   0  - no problems, or problems found but -FailOnZombie not requested
#   1  - -FailOnZombie was set and at least one zombie was detected

[CmdletBinding()]
param(
    [int]$AgeSeconds = 120,
    [switch]$FailOnZombie
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

# Reconcile first so live processes become active and dead active entries move
# to inactive (the exact signal this script then analyzes).
try {
    Sync-ExcelInstanceSnapshot | Out-Null
} catch {
    Write-Warning "Sync failed (continuing with stale registry): $($_.Exception.Message)"
}

$active = @(Read-ExcelInstances)
$inactive = @(Read-InactiveExcelInstances)
$running = @(Get-RunningExcelInstances)

$livePids = @{}
foreach ($p in $running) {
    $livePids[[int]$p.pid] = $true
}

$now = Get-Date

# 1) Zombies: deactivated entries whose process is still running.
$zombies = @($inactive | Where-Object { $livePids.ContainsKey([int]$_.pid) })

# 2) Lingering: live automation-owned instances older than the threshold.
$lingering = @()
foreach ($entry in $active) {
    $instancePid = [int]$entry.pid
    if (-not $livePids.ContainsKey($instancePid)) { continue }

    # Only flag hidden (automation-created) instances. A visible instance is a
    # user session (or the reusable visible add-in host opened by globalSetup)
    # and is intentionally left open, so it must never count as "lingering".
    if ([bool]$entry.visible) { continue }

    $age = 0
    try { $age = ($now - [datetime]::Parse([string]$entry.createdAt)).TotalSeconds } catch {}
    if ($age -ge $AgeSeconds) {
        $lingering += [pscustomobject]@{
            pid        = $instancePid
            ageSeconds = [math]::Floor($age)
            entry      = $entry
        }
    }
}

# Report (and collect log lines for the shared instance log).
$logLines = @()

$summary = "Assess-ExcelInstances: active=$($active.Count) inactive=$($inactive.Count) running=$($running.Count) zombies=$($zombies.Count) lingering=$($lingering.Count)"
Write-Output $summary
$logLines += "[vbapm-assess] $((Get-Date).ToString('o')) $summary"

function Format-EntryWorkbooks {
    param([object]$Entry)
    $wbs = @($Entry.workbooks)
    if ($wbs.Count -gt 0 -and $null -ne $wbs[0]) {
        return ($wbs -join ', ')
    }
    return '(none recorded)'
}

foreach ($z in $zombies) {
    $wbStr = Format-EntryWorkbooks $z
    $msg = "ZOMBIE pid=$($z.pid) owner=$($z.owner) reason=$($z.reason) deactivatedAt=$($z.deactivatedAt) deactivateReason=$($z.deactivateReason) workbooks=$wbStr"
    Write-Warning $msg
    $logLines += "[vbapm-assess] $((Get-Date).ToString('o')) $msg"

    $wbs = @($z.workbooks)
    if ($wbs.Count -gt 0 -and $null -ne $wbs[0]) {
        foreach ($w in $wbs) {
            Write-Warning "          workbook: $w"
        }
    }
}

foreach ($l in $lingering) {
    $e = $l.entry
    $wbStr = Format-EntryWorkbooks $e
    $msg = "LINGERING pid=$($l.pid) age=$($l.ageSeconds)s owner=$($e.owner) reason=$($e.reason) workbooks=$wbStr"
    Write-Warning $msg
    $logLines += "[vbapm-assess] $((Get-Date).ToString('o')) $msg"

    $wbs = @($e.workbooks)
    if ($wbs.Count -gt 0 -and $null -ne $wbs[0]) {
        foreach ($w in $wbs) {
            Write-Warning "          workbook: $w"
        }
    }
}

# Append findings to the shared instance log for post-run review.
try {
    $logPath = if ($env:VBA_INSTANCE_LOG) { $env:VBA_INSTANCE_LOG } else { Join-Path $env:TEMP 'Excel-Instances\instances.log' }
    Add-Content -LiteralPath $logPath -Value $logLines -ErrorAction Stop
} catch {
    # best-effort; never fail the assessment over a log write.
}

if ($zombies.Count -eq 0 -and $lingering.Count -eq 0) {
    Write-Output "No lingering or zombie Excel instances detected."
    exit 0
}

if ($FailOnZombie -and $zombies.Count -gt 0) {
    Write-Output "FAIL: $($zombies.Count) zombie Excel instance(s) detected."
    exit 1
}

exit 0
