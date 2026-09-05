# Make-ExcelVisible.ps1
#
# Show every top-level window owned by one Excel process.

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, Position = 0)]
    [Alias('Pid')]
    [int]$ProcessId
)

$ErrorActionPreference = 'Stop'

$helper = Join-Path $PSScriptRoot 'Set-ExcelWindowsVisibility.ps1'
& $helper -ProcessId $ProcessId -Visible:$true -WhatIf:$WhatIfPreference
exit $LASTEXITCODE