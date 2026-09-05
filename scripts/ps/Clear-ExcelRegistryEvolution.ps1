# Clear-ExcelRegistryEvolution.ps1
#
# Remove all prior registry evolution snapshots before a new evolution run.

[CmdletBinding(SupportsShouldProcess)]
param()

$ErrorActionPreference = 'Stop'
$evolutionDir = Join-Path $env:TEMP 'Excel-Instances\evolution'

if (-not (Test-Path -LiteralPath $evolutionDir)) {
    exit 0
}

foreach ($item in @(Get-ChildItem -LiteralPath $evolutionDir -Force)) {
    if ($PSCmdlet.ShouldProcess($item.FullName, 'Remove evolution artifact')) {
        Remove-Item -LiteralPath $item.FullName -Recurse -Force
    }
}
