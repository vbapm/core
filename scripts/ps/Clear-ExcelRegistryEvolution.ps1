# Clear-ExcelRegistryEvolution.ps1
#
# Start a new registry evolution run without removing prior runs.

[CmdletBinding(SupportsShouldProcess)]
param()

$ErrorActionPreference = 'Stop'
$evolutionDir = Join-Path $env:TEMP 'Excel-Instances\evolution'
New-Item -ItemType Directory -Path $evolutionDir -Force | Out-Null

$stamp = [DateTime]::UtcNow.ToString("yyyy-MM-dd--HH'h'mm'm'ss's'fff")
$runDir = Join-Path $evolutionDir $stamp
$suffix = 1
while (Test-Path -LiteralPath $runDir) {
    $runDir = Join-Path $evolutionDir "$stamp-$suffix"
    $suffix++
}

New-Item -ItemType Directory -Path $runDir -Force | Out-Null
$marker = Join-Path $evolutionDir 'current-run.txt'
[System.IO.File]::WriteAllText($marker, $runDir, [System.Text.UTF8Encoding]::new($false))
Write-Output "Evolution run directory: $runDir"
