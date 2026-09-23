# Register-ExcelInstance.ps1
#
# Manually track a running Excel.Application instance in the coordination
# registry. Useful for agents that create an Excel instance outside of the
# run.ps1 bridge (e.g. a one-off COM automation script).
#
# Usage:
#   .\scripts\ps\Register-ExcelInstance.ps1 -ProcessId 1234 -Reason e2e
#   .\scripts\ps\Register-ExcelInstance.ps1 -ProcessId 1234 -Reason manual -Visible

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [int]$ProcessId,

    [string]$Owner,

    [string]$Reason = 'manual',

    [switch]$Visible
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

if (-not $Owner) {
    # Trace ownership back to this terminal/process. The register script runs
    # *inside* the calling terminal, so $PID is that terminal's process id.
    $Owner = "terminal-$PID"
}

Update-ExcelInstance -InstanceId $ProcessId -Owner $Owner -Visible ([bool]$Visible) -Reason $Reason

Write-Output "Registered Excel instance pid=$ProcessId owner=$Owner reason=$Reason visible=$Visible"
