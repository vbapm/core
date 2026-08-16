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
#
# When -FailOnRogue is set, the script exits non-zero if any EXCEL.EXE process
# exists that is not present in the registry. This is the guard recommended to
# run before starting an e2e integration suite so we don't hijack another
# agent's or a user's Excel session.

[CmdletBinding()]
param(
    [switch]$Json,
    [switch]$FailOnRogue,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

$registered = @(Read-ExcelInstances)
$running = @(Get-RunningExcelInstances)

$registeredPids = @{}
foreach ($r in $registered) {
    $registeredPids[[int]$r.pid] = $true
}

$rogue = @($running | Where-Object { -not $registeredPids.ContainsKey([int]$_.pid) })
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
    runningCount    = $running.Count
    rogueCount      = $rogue.Count
    orphanedCount   = $orphaned.Count
    registered      = @($registered)
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
    Write-Output ("Running EXCEL count: {0}" -f $status.runningCount)
    Write-Output ("Rogue count        : {0}" -f $status.rogueCount)
    Write-Output ("Orphaned count     : {0}" -f $status.orphanedCount)
    Write-Output ""

    if ($status.registeredCount -gt 0) {
        Write-Output "--- Registered instances ---"
        foreach ($r in $registered) {
            $idStr = if ($r.id) { $r.id } else { '(none)' }
            $comStr = if ($r.comReachable) { 'com' } else { 'no-com' }
            Write-Output ("  id={0,-10} pid={1,-6} owner={2} reason={3} visible={4} com={5} createdAt={6}" -f $idStr, $r.pid, $r.owner, $r.reason, $r.visible, $comStr, $r.createdAt)
            if ($r.windowTitle) {
                Write-Output ("          window: {0}" -f $r.windowTitle)
            }
            $wbs = @($r.workbooks)
            if ($wbs.Count -gt 0) {
                foreach ($w in $wbs) {
                    Write-Output ("          workbook: {0}" -f $w)
                }
            }
            $addins = @($r.addins)
            if ($addins.Count -gt 0) {
                foreach ($a in $addins) {
                    $openFlag = if ($a.isOpen) { 'open' } else { 'closed' }
                    Write-Output ("          addin: {0} [{1}]" -f $a.name, $openFlag)
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
