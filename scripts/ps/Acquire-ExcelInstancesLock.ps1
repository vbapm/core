# Acquire-ExcelInstancesLock.ps1
#
# Acquire the Excel instance registry lock and hold it until the caller
# releases it. Intended for agents that need to atomically read-modify-write the
# registry across several operations they don't want interleaved with another
# agent. The lock is released automatically when this script exits via a
# Process exit, or manually by running the -Release switch.
#
# Usage:
#   .\scripts\ps\Acquire-ExcelInstancesLock.ps1        # acquire, then block? No:
#                                                       # prints a token and releases at end
#   .\scripts\ps\Acquire-ExcelInstancesLock.ps1 -Release

[CmdletBinding()]
param(
    [switch]$Release
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

if ($Release) {
    Release-ExcelInstancesLock
    Write-Output "Lock released."
    exit 0
}

Get-ExcelInstancesLock | Out-Null
Write-Output "Lock acquired (held until this process exits or release is called)."
Write-Output "Lock file: $(Get-ExcelInstancesLockPath)"
