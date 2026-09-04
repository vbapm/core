# Close-ExcelInstance.ps1
#
# Close a tracked Excel instance by its stable hash `id` (preferred) or by
# process id. After a successful close, the instance is removed from the
# coordination registry and a resync is triggered (closing an instance can
# change which remaining instance is COM-reachable via the ROT).
#
# Usage:
#   .\scripts\ps\Close-ExcelInstance.ps1 -Id a1b2c3d4
#   .\scripts\ps\Close-ExcelInstance.ps1 -ProcessId 1234

[CmdletBinding(DefaultParameterSetName = 'ById')]
param(
    [Parameter(Mandatory, ParameterSetName = 'ById')]
    [string]$Id,

    [Parameter(Mandatory, ParameterSetName = 'ByPid')]
    [int]$ProcessId
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

$closed = $false
$closeMessage = ''

if ($PSCmdlet.ParameterSetName -eq 'ById') {
    $entry = Find-ExcelInstanceById -Id $Id
    if (-not $entry) {
        Write-Output "No Excel instance with id=$Id found in the registry."
        exit 1
    }
    $closed = Close-ExcelInstanceById -Id $Id
    $closeMessage = "Closed Excel instance id=$Id pid=$($entry.pid) (killed=$closed)"
} else {
    # For PID-based close, kill the process then remove it from the registry.
    $proc = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if ($proc -and $proc.ProcessName -eq 'EXCEL') {
        try { Stop-Process -Id $ProcessId -Force -ErrorAction Stop; $closed = $true } catch { $closed = $false }
    } else {
        $closed = $false
    }
    Remove-ExcelInstance -InstanceId $ProcessId
    $closeMessage = "Closed Excel instance pid=$ProcessId (killed=$closed)"
}

Write-Output $closeMessage

# Trigger a resync so the registry reflects the new COM-reachability landscape
# (e.g. a previously-unreachable instance may now own the ROT slot).
$syncScript = Join-Path $PSScriptRoot 'Sync-ExcelInstances.ps1'
if (Test-Path -LiteralPath $syncScript) {
    Write-Output 'Resyncing instance registry...'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript -PruneDead
}

