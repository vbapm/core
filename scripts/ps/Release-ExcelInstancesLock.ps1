# Release-ExcelInstancesLock.ps1
#
# Explicitly release the Excel instance registry lock (in case it was acquired
# out-of-band or a previous holder crashed without cleanup).
#
# Usage:
#   .\scripts\ps\Release-ExcelInstancesLock.ps1

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

Release-ExcelInstancesLock
Write-Output "Lock released (if held)."
