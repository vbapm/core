# Make-ExcelInvisible.ps1
#
# Hide every top-level window owned by one Excel process.
# Use -nb to select the active registry number shown by List-ExcelInstances.ps1.
# Examples: .\Make-ExcelInvisible.ps1 -nb 2 or .\Make-ExcelInvisible.ps1 26368

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, Position = 0, ParameterSetName = 'ProcessId')]
    [Alias('Pid')]
    [int]$ProcessId,

    [Parameter(Mandatory, ParameterSetName = 'Number')]
    [Alias('nb')]
    [int]$Number
)

$ErrorActionPreference = 'Stop'

$processIdToHide = $ProcessId
if ($PSCmdlet.ParameterSetName -eq 'Number') {
    . (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')
    $entry = Find-ExcelInstanceByNumber -Number $Number
    if ($null -eq $entry) {
        throw "No active Excel instance has registry number $Number."
    }

    $processIdToHide = [int]$entry.pid
}

$helper = Join-Path $PSScriptRoot 'Set-ExcelWindowsVisibility.ps1'
& $helper -ProcessId $processIdToHide -Visible:$false -WhatIf:$WhatIfPreference
exit $LASTEXITCODE
