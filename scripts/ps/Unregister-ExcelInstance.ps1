# Unregister-ExcelInstance.ps1
#
# Remove an Excel instance from the coordination registry. Call this when an
# instance is being torn down so the registry doesn't go stale.
#
# Usage:
#   .\scripts\ps\Unregister-ExcelInstance.ps1 -ProcessId 1234

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [int]$ProcessId
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

Remove-ExcelInstance -InstanceId $ProcessId

Write-Output "Unregistered Excel instance pid=$ProcessId"
