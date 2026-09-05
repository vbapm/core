# Close-AllInvisibleExcelInstances.ps1
#
# Close every invisible Excel instance. Visible user sessions are left alone.

[CmdletBinding()]
param(
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

$helper = Join-Path $PSScriptRoot 'Close-ExcelInstances.ps1'
& $helper -Scope Invisible -WhatIf:$WhatIf
exit 0
