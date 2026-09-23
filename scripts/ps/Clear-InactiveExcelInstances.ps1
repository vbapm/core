# Clear-InactiveExcelInstances.ps1
#
# Reset the registry's "inactive" (recently deactivated) instance list. Run at
# the start of an e2e suite so the end-of-suite assessment only considers
# instances deactivated during the current run (and doesn't report zombies from
# previous runs).
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\scripts\ps\Clear-InactiveExcelInstances.ps1

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

Clear-InactiveExcelInstances

Write-Output "Cleared inactive Excel instances."
