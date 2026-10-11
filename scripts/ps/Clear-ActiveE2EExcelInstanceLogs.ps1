# Clear logs on active hidden e2e Excel instances before a new suite starts.
# Historical inactive records and visible/user instances are preserved.

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')
Clear-ActiveE2EExcelInstanceLogs
Write-Output 'Cleared active hidden e2e Excel instance logs.'
