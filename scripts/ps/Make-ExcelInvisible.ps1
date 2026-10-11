# Make-ExcelInvisible.ps1
#
# Hide one Excel instance selected by its instances.json number.
# Use -nb to select the active registry number shown by List-ExcelInstances.ps1.
# Examples: .\Make-ExcelInvisible.ps1 1 or .\Make-ExcelInvisible.ps1 -ProcessId 26368

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, ParameterSetName = 'ProcessId')]
    [Alias('Pid')]
    [int]$ProcessId,

    [Parameter(Mandatory, Position = 0, ParameterSetName = 'Number')]
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
