# Make-ExcelInvisible.ps1
#
# Hide every top-level window owned by one Excel process.

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, Position = 0)]
    [Alias('Pid')]
    [int]$ProcessId
)

$ErrorActionPreference = 'Stop'

$helper = Join-Path $PSScriptRoot 'Set-ExcelWindowsVisibility.ps1'
& $helper -ProcessId $ProcessId -Visible:$false -WhatIf:$WhatIfPreference
exit $LASTEXITCODE
