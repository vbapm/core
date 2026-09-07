# Get-ExcelInstancesStatus.ps1
#
# Reports the coordination registry plus live EXCEL.EXE processes, so agents
# can tell which Excel instances are tracked by the registry (i.e. created by
# automation) versus which are "rogue" — running but unrecorded (a user session
# or an agent that failed to clean up).
#
# Usage:
#   .\scripts\ps\Get-ExcelInstancesStatus.ps1
#   .\scripts\ps\Get-ExcelInstancesStatus.ps1 -Json            # machine-readable
#   .\scripts\ps\Get-ExcelInstancesStatus.ps1 -FailOnRogue      # exit 1 if rogue
#   .\scripts\ps\Get-ExcelInstancesStatus.ps1 -TrackRogue      # persist new rogue entries
#
# When -FailOnRogue is set, the script exits non-zero if any EXCEL.EXE process
# exists that is not present in the registry. This is the guard recommended to
# run before starting an e2e integration suite so we don't hijack another
# agent's or a user's Excel session.

[CmdletBinding()]
param(
    [switch]$Json,
    [switch]$FailOnRogue,
    [switch]$TrackRogue,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

# Reconcile first so the report reflects *all* live EXCEL instances. New
# process-only rogues are persisted only when -TrackRogue is requested; the
# default keeps them untracked while still reporting them live.
try {
    Sync-ExcelInstanceSnapshot -TrackRogue:$TrackRogue | Out-Null
} catch {
    # Report is best-effort; never fail the status command over a sync error.
}

$registered = @(Read-ExcelInstances)
$inactive = @(Read-InactiveExcelInstances)
$running = @(Get-RunningExcelInstances)

$registeredPids = @{}
foreach ($r in $registered) {
    $registeredPids[[int]$r.pid] = $true
}

$rogue = @($running | Where-Object {
    $process = $_
    $entry = @($registered | Where-Object { [int]$_.pid -eq [int]$process.pid })[0]
    -not $registeredPids.ContainsKey([int]$process.pid) -or
        ($entry -and [string]$entry.owner -match '^(rogue|unknown|unknown-owner|untracked)$')
})
$orphaned = @($registered | Where-Object {
    $instancePid = [int]$_.pid
    -not ($running | Where-Object { [int]$_.pid -eq $instancePid })
})

$status = [pscustomobject]@{
    directory       = (Get-ExcelInstancesDir)
    registryPath    = (Get-ExcelInstancesPath)
    lockPath        = (Get-ExcelInstancesLockPath)
    lockHeld        = [bool](Test-Path -LiteralPath (Get-ExcelInstancesLockPath))
    registeredCount = $registered.Count
    inactiveCount   = $inactive.Count
    runningCount    = $running.Count
    rogueCount      = $rogue.Count
    invisibleCount  = @($running | Where-Object { -not $_.visible }).Count
    orphanedCount   = $orphaned.Count
    registered      = @($registered)
    inactive        = @($inactive)
    running         = @($running)
    rogue           = @($rogue)
    orphaned        = @($orphaned)
}

if ($Json) {
    $status | ConvertTo-Json -Depth 5
} elseif (-not $Quiet) {
    Write-Output "=== Excel Instance Registry Status ==="
    Write-Output ("Registry directory : {0}" -f $status.directory)
    Write-Output ("Lock file held     : {0}" -f $status.lockHeld)
    Write-Output ("Registered count   : {0}" -f $status.registeredCount)
    Write-Output ("Inactive count     : {0}" -f $status.inactiveCount)
    Write-Output ("Running EXCEL count: {0}" -f $status.runningCount)
    Write-Output ("Rogue count        : {0}" -f $status.rogueCount)
    Write-Output ("Invisible count    : {0}" -f $status.invisibleCount)
    Write-Output ("Orphaned count     : {0}" -f $status.orphanedCount)
    Write-Output ""

    if ($status.registeredCount -gt 0) {
        Write-Output "--- Instances ---"
        foreach ($r in $registered) {
            $idStr = if ($r.id) { $r.id } else { '(none)' }
            $comStr = if ($r.comReachable) { 'com' } else { 'no-com' }
            Write-Output ("  id={0,-10} pid={1,-6} owner={2} reason={3} visible={4} com={5} createdAt={6}" -f $idStr, $r.pid, $r.owner, $r.reason, $r.visible, $comStr, $r.createdAt)
            if ($r.windowTitle) {
                Write-Output ("          window: {0}" -f $r.windowTitle)
            }
            $wbs = @($r.workbooks)
            if ($wbs.Count -gt 0 -and $null -ne $wbs[0]) {
                foreach ($w in $wbs) {
                    Write-Output ("          workbook: {0}" -f $w)
                }
            }
            $addins = @($r.addins)
            if ($addins.Count -gt 0 -and $null -ne $addins[0]) {
                foreach ($a in $addins) {
                    $openFlag = if ($null -eq $a.isOpen) { 'unknown' } elseif ($a.isOpen) { 'open' } else { 'closed' }
                    $installedFlag = if ($null -eq $a.isInstalled) { '?' } elseif ($a.isInstalled) { 'installed' } else { 'not-installed' }
                    Write-Output ("          addin: {0} [isInstalled={1} {2}]" -f $a.name, $installedFlag, $openFlag)
                }
            }
        }
        Write-Output ""
    }

    if ($status.inactiveCount -gt 0) {
        Write-Output "--- Recently deactivated (inactive) ---"
        foreach ($r in $inactive) {
            $idStr = if ($r.id) { $r.id } else { '(none)' }
            Write-Output ("  id={0,-10} pid={1,-6} owner={2} reason={3} deactivatedAt={4} why={5}" -f $idStr, $r.pid, $r.owner, $r.reason, $r.deactivatedAt, $r.deactivateReason)
            $wbs = @($r.workbooks)
            if ($wbs.Count -gt 0 -and $null -ne $wbs[0]) {
                foreach ($w in $wbs) {
                    Write-Output ("          workbook: {0}" -f $w)
                }
            }
        }
        Write-Output ""
    }

    if ($status.runningCount -gt 0) {
        Write-Output "--- Live EXCEL.EXE processes ---"
        foreach ($p in $running) {
            $tag = if ($registeredPids.ContainsKey([int]$p.pid)) { 'tracked' } else { 'ROGUE' }
            Write-Output ("  pid={0,-6} visible={1,-5} {2,-7} title={3}" -f $p.pid, $p.visible, $tag, $p.mainWindowTitle)
        }
        Write-Output ""
    }

    if ($status.rogueCount -gt 0) {
        Write-Warning "$($status.rogueCount) Excel process(es) are running but NOT tracked in the registry:"
        foreach ($p in $rogue) {
            Write-Warning "  ROGUE pid=$($p.pid) title=$($p.mainWindowTitle)"
        }
    }

    if ($status.orphanedCount -gt 0) {
        Write-Warning "$($status.orphanedCount) registry entr(ies) point to dead Excel processes (stale):"
        foreach ($o in $orphaned) {
            Write-Warning "  ORPHANED pid=$($o.pid) owner=$($o.owner)"
        }
    }
}

if ($FailOnRogue -and $status.rogueCount -gt 0) {
    Write-Output "FAIL: rogue Excel instances detected. Aborting to avoid interfering with another session."
    exit 1
}

exit 0
