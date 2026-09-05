# Close-AllExcelInstances.ps1
#
# Close every running Excel instance, visible or invisible. Use this carefully
# because it can close a user's work.

[CmdletBinding()]
param(
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

$helper = Join-Path $PSScriptRoot 'Close-ExcelInstances.ps1'
& $helper -Scope All -WhatIf:$WhatIf
exit 0
