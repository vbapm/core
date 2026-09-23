# Close-AllInvisibleExcelInstances.ps1
#
# Kill every *invisible* Excel instance (hidden window, i.e. automation-created),
# whether it is tracked in the coordination registry or has leaked ("rogue").
# Visible user sessions are left untouched.
#
# This is the cleanup used to reclaim instances leaked by aborted/crashed e2e
# runs. It is intentionally conservative: only hidden instances are killed.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\scripts\ps\Close-AllInvisibleExcelInstances.ps1
#   powershell -ExecutionPolicy Bypass -File .\scripts\ps\Close-AllInvisibleExcelInstances.ps1 -WhatIf

[CmdletBinding()]
param(
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

$running = @(Get-RunningExcelInstances)
$invisible = @($running | Where-Object { -not $_.visible })

if ($invisible.Count -eq 0) {
    Write-Output "No invisible Excel instances found."
    exit 0
}

Write-Output "Closing $($invisible.Count) invisible Excel instance(s)..."
$killed = 0
foreach ($p in $invisible) {
    if ($WhatIf) {
        Write-Output "  [WhatIf] would kill pid=$($p.pid)"
        continue
    }
    try {
        Stop-Process -Id $p.pid -Force -ErrorAction Stop
        Write-Output "  killed pid=$($p.pid)"
        $killed++
    } catch {
        Write-Output "  failed pid=$($p.pid): $($_.Exception.Message)"
    }

    # Also drop any registry entry for it so the registry doesn't go stale.
    try { Remove-ExcelInstance -InstanceId $p.pid } catch {}
}

Write-Output "Done. Killed $killed invisible instance(s)."

# Resync the registry to prune any now-dead entries.
$syncScript = Join-Path $PSScriptRoot 'Sync-ExcelInstances.ps1'
if (Test-Path -LiteralPath $syncScript) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript -PruneDead
}

exit 0
