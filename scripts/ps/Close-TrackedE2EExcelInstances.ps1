# Close-TrackedE2EExcelInstances.ps1
#
# Final e2e cleanup for hidden Excel instances owned by persistent test sessions.
# Visible Excel and untracked/rogue instances are intentionally left alone.

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

$tracked = @(Read-ExcelInstances | Where-Object {
    (-not [bool]$_.visible) -and (
        [string]$_.reason -eq 'e2e' -or
        [string]$_.owner -like 'session#*'
    )
})

foreach ($entry in $tracked) {
    $instancePid = [int]$entry.pid
    try {
        Stop-Process -Id $instancePid -Force -ErrorAction SilentlyContinue
    } catch {}

    try {
        Remove-ExcelInstance -InstanceId $instancePid -DeactivateReason 'e2e-final-cleanup'
    } catch {}
}

Write-Output "Closed $($tracked.Count) tracked e2e Excel instance(s)."
