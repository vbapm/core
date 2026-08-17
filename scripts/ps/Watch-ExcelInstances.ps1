# Watch-ExcelInstances.ps1
#
# Debugging aid for leaked Excel instances. Scans live EXCEL.EXE processes and
# reports any that have been running longer than a threshold, optionally making
# invisible (automation) instances visible so a stuck instance can be seen and
# diagnosed instead of remaining a hidden zombie.
#
# Making a hidden instance visible is best-effort: hidden automation instances
# have MainWindowHandle == 0 and may not be registered in the Running Object
# Table, so we try (1) the ROT (Hwnd -> PID match) and fall back to (2) reporting
# the instance as unreachable.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\scripts\ps\Watch-ExcelInstances.ps1
#   powershell -ExecutionPolicy Bypass -File .\scripts\ps\Watch-ExcelInstances.ps1 -AgeSeconds 120
#   powershell -ExecutionPolicy Bypass -File .\scripts\ps\Watch-ExcelInstances.ps1 -AgeSeconds 120 -MakeVisible

[CmdletBinding()]
param(
    [int]$AgeSeconds = 120,
    [switch]$MakeVisible
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

$now = Get-Date
$running = @(Get-RunningExcelInstances)
$stale = @($running | Where-Object {
    $start = [datetime]::Parse($_.startedAt)
    ($now - $start).TotalSeconds -ge $AgeSeconds
})

if ($stale.Count -eq 0) {
    Write-Output "No Excel instance older than $AgeSeconds seconds."
    exit 0
}

Write-Output "Found $($stale.Count) Excel instance(s) older than $AgeSeconds seconds:"
foreach ($p in $stale) {
    $age = [math]::Floor(($now - [datetime]::Parse($p.startedAt)).TotalSeconds)
    Write-Output ("  pid={0,-6} visible={1,-5} age={2}s title={3}" -f $p.pid, $p.visible, $age, $p.mainWindowTitle)
}

if (-not $MakeVisible) {
    Write-Output "Re-run with -MakeVisible to force invisible instances visible."
    exit 0
}

Write-Output ""
Write-Output "Attempting to make invisible stale instances visible..."
$madeVisible = 0

foreach ($p in $stale) {
    if ($p.visible) { continue }

    $app = $null
    # Try to locate the COM Application whose Hwnd maps to this PID via the ROT.
    foreach ($a in @(Get-ExcelApplications)) {
        try {
            if ((Get-ExcelProcessId -ExcelApp $a) -eq $p.pid) {
                $app = $a
                break
            }
        } catch {
            # keep scanning
        }
    }

    if ($null -ne $app) {
        try {
            $app.Visible = $true
            Write-Output "  pid=$($p.pid) made visible."
            $madeVisible++
            continue
        } catch {
            Write-Warning "  pid=$($p.pid) found in ROT but failed to set Visible: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "  pid=$($p.pid) not reachable via ROT (hidden instance); cannot make visible. Consider Close-AllInvisibleExcelInstances.ps1."
    }
}

Write-Output "Done. Made $madeVisible stale instance(s) visible."

exit 0
