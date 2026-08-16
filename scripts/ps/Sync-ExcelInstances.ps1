# Sync-ExcelInstances.ps1
#
# Enumerate every running Excel.Application instance on this machine (via the
# Running Object Table) and (re)write the coordination registry with up-to-date
# visibility and open-workbook information for each. Instances that were already
# tracked keep their existing `owner` / `reason`; newly discovered instances are
# recorded with a `manual` owner and `manual` reason.
#
# This is how a user or agent can register an already-open (e.g. visible) Excel
# session into %TEMP%\Excel-Instances\instances.json without going through the
# run.ps1 bridge.
#
# Usage:
#   .\scripts\ps\Sync-ExcelInstances.ps1
#   .\scripts\ps\Sync-ExcelInstances.ps1 -PruneDead    # also drop stale entries

[CmdletBinding()]
param(
    [switch]$PruneDead
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

<#
.SYNOPSIS
Resolve a meaningful owner label for an Excel process when one isn't already
tracked. COM launches Excel via svchost.exe (the COM SCM), which breaks the
process-parent chain, so we can't always recover the originating terminal.
Fall back to a descriptive label rather than "manual".
#>
function Resolve-ExcelInstanceOwner {
    param([int]$InstanceId)

    try {
        $proc = Get-CimInstance Win32_Process -Filter "ProcessId = $InstanceId" -ErrorAction Stop
        $parent = Get-CimInstance Win32_Process -Filter "ProcessId = $($proc.ParentProcessId)" -ErrorAction SilentlyContinue
        if ($parent -and $parent.Name -notmatch '^(svchost|services)\.exe$') {
            return "terminal-$($parent.ProcessId)"
        }
    } catch {
        # fall through
    }
    return 'unknown-owner'
}

<#
.SYNOPSIS
Helper: write an instance entry, preserving prior owner/reason when the PID was
already tracked.
#>
function Upsert-Instance {
    param(
        [int]$InstanceId,
        [bool]$Visible,
        [string[]]$Workbooks,
        [bool]$ComReachable = $false
    )

    $owner = Resolve-ExcelInstanceOwner $InstanceId
    $reason = 'unknown'
    if ($existingByPid.ContainsKey([int]$InstanceId)) {
        $prev = $existingByPid[[int]$InstanceId]
        $owner = if ($prev.owner) { $prev.owner } else { $owner }
        $reason = if ($prev.reason) { $prev.reason } else { $reason }
        # Preserve a prior comReachable=true if this PID was ever COM-reachable,
        # unless it is explicitly being re-observed as not reachable now.
        if (-not $ComReachable -and $prev.comReachable -eq $true) {
            $ComReachable = $true
        }
    }

    Update-ExcelInstance -InstanceId $InstanceId -Owner $owner -Visible $Visible -Reason $reason -Workbooks $Workbooks -ComReachable $ComReachable
}

$apps = @(Get-ExcelApplications)

# Snapshot existing entries to preserve owner/reason for already-tracked PIDs.
$existing = @(Read-ExcelInstances)
$existingByPid = @{}
foreach ($e in $existing) {
    $existingByPid[[int]$e.pid] = $e
}

$livePids = @{}

# 1) Sync every COM-reachable instance (via ROT) with full visibility + workbooks.
foreach ($app in $apps) {
    try {
        $instancePid = Get-ExcelProcessId -ExcelApp $app
    } catch {
        Write-Warning "Could not resolve PID for a ROT Excel instance; skipping."
        continue
    }
    $livePids[[int]$instancePid] = $true

    $isVisible = $false
    try { $isVisible = [bool]$app.Visible } catch {}

    $openWorkbooks = @(Get-ExcelWorkbooks -ExcelApp $app)

    Upsert-Instance $instancePid $isVisible $openWorkbooks $true
    Write-Output "Synced (ROT) pid=$instancePid visible=$isVisible workbooks=$($openWorkbooks.Count)"
}

# 2) Discover any EXCEL process NOT reachable via ROT (e.g. a user-visible
#    instance launched by file association). For these we can still record
#    visibility and — best-effort — the workbook names visible in window titles.
$allExcel = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
foreach ($p in $allExcel) {
    $instancePid = [int]$p.Id
    if ($livePids.ContainsKey($instancePid)) { continue }

    $livePids[$instancePid] = $true

    $isVisible = [bool]($p.MainWindowHandle -ne 0)

    # Best-effort workbook list from this process's visible main window title.
    # Format is "<WorkbookName> - Excel". Fall back to placeholder when unknown.
    $openWorkbooks = @()
    if ($isVisible -and $p.MainWindowTitle) {
        $title = $p.MainWindowTitle
        if ($title -match '^(.*?)(\.xlsx|\.xlsm|\.xlsb|\.xls)?\s*-\s*Excel$') {
            $openWorkbooks += $p.MainWindowTitle -replace '\s*-\s*Excel$', ''
        } else {
            $openWorkbooks += "[$($p.MainWindowTitle)]"
        }
    } else {
        $openWorkbooks += '[unknown — instance not reachable via COM]'
    }

    Upsert-Instance $instancePid $isVisible $openWorkbooks $false
    Write-Output "Synced (process) pid=$instancePid visible=$isVisible workbooks=$($openWorkbooks.Count)"
}

if ($livePids.Count -eq 0) {
    Write-Output "No running Excel instances found."
}

# Optionally drop registry entries whose PIDs are no longer live.
if ($PruneDead) {
    Get-ExcelInstancesLock | Out-Null
    try {
        $all = @(Read-ExcelInstances)
        $kept = @($all | Where-Object { $livePids.ContainsKey([int]$_.pid) })
        if ($kept.Count -ne $all.Count) {
            Write-ExcelInstances $kept
            Write-Output "Pruned $($all.Count - $kept.Count) stale entr(ies)."
        }
    } finally {
        Release-ExcelInstancesLock
    }
}
