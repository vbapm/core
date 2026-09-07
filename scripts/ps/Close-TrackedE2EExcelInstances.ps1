# Close-TrackedE2EExcelInstances.ps1
#
# Final e2e cleanup for hidden Excel instances owned by persistent test sessions.
# Visible Excel and untracked/rogue instances are intentionally left alone.

[CmdletBinding()]
param(
    [int]$KeepWarmCount = 0,
    [switch]$KeepOneWarm
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

$tracked = @(Read-ExcelInstances | Where-Object {
    (-not [bool]$_.visible) -and (
        [string]$_.reason -eq 'e2e' -or
        [string]$_.owner -like 'session#*'
    )
})

$warmCount = [Math]::Max(0, $KeepWarmCount)
if ($KeepOneWarm) {
    $warmCount = [Math]::Max($warmCount, 1)
}
$warmPids = @($tracked | Sort-Object createdAt | Select-Object -First $warmCount | ForEach-Object { [int]$_.pid })

$closed = 0
if ($warmPids.Count -gt 0) {
    Write-Output "Keeping tracked e2e Excel instance(s) warm: $($warmPids -join ', ')."
}

foreach ($entry in $tracked) {
    $instancePid = [int]$entry.pid
    if ($warmPids -contains $instancePid) { continue }
    try {
        Stop-Process -Id $instancePid -Force -ErrorAction SilentlyContinue
    } catch {}

    try {
        Remove-ExcelInstance -InstanceId $instancePid -DeactivateReason 'e2e-final-cleanup'
    } catch {}
    $closed++
}

Write-Output "Closed $closed tracked e2e Excel instance(s)."
