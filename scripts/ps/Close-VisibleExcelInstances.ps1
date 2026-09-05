# Close-VisibleExcelInstances.ps1
#
# Close every visible Excel instance. Use this carefully because it can close
# a user's work.

[CmdletBinding()]
param(
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

$helper = Join-Path $PSScriptRoot 'Close-ExcelInstances.ps1'
& $helper -Scope Visible -WhatIf:$WhatIf
exit 0
