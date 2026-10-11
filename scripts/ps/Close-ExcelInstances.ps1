# Close-ExcelInstances.ps1
#
# Shared implementation for closing Excel instances selected by visibility.
# The public wrappers are Close-AllInvisibleExcelInstances.ps1,
# Close-VisibleExcelInstances.ps1, and Close-AllExcelInstances.ps1.

[CmdletBinding()]
param(
    [ValidateSet('Invisible', 'Visible', 'All')]
    [string]$Scope = 'Invisible',
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

$running = @(Get-RunningExcelInstances)
switch ($Scope) {
    'Invisible' {
        $targets = @($running | Where-Object { -not $_.visible })
        $label = 'invisible'
    }
    'Visible' {
        $targets = @($running | Where-Object { $_.visible })
        $label = 'visible'
    }
    'All' {
        $targets = $running
        $label = ''
    }
}

if ($targets.Count -eq 0) {
    if ($label) {
        Write-Output "No $label Excel instances found."
    } else {
        Write-Output 'No Excel instances found.'
    }
    exit 0
}

$description = if ($label) { "$label Excel" } else { 'Excel' }
Write-Output "Closing $($targets.Count) $description instance(s)..."
$killed = 0
foreach ($instance in $targets) {
    if ($WhatIf) {
        Write-Output "  [WhatIf] would kill pid=$($instance.pid)"
        continue
    }

    try {
        Stop-Process -Id $instance.pid -Force -ErrorAction Stop
        Write-Output "  killed pid=$($instance.pid)"
        $killed++
    } catch {
        Write-Output "  failed pid=$($instance.pid): $($_.Exception.Message)"
    }

    try { Remove-ExcelInstance -InstanceId $instance.pid } catch {}
}

if ($WhatIf) {
    Write-Output 'Done. No processes were changed.'
    exit 0
}

Write-Output "Done. Killed $killed $description instance(s)."

$syncScript = Join-Path $PSScriptRoot 'Sync-ExcelInstances.ps1'
if (Test-Path -LiteralPath $syncScript) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript -PruneDead
}

exit 0
