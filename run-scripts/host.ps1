[CmdletBinding()]
param(
    [ValidateSet('list', 'close-all', 'close-visible', 'close-invisible')]
    [string]$Action = 'list'
)

$ErrorActionPreference = 'Stop'

function Get-ExcelInstanceInfo {
    param([Parameter(Mandatory)][System.Diagnostics.Process]$Process)

    $handle = 0
    $title = ''
    $startedAt = $null
    try { $handle = [int64]$Process.MainWindowHandle } catch {}
    try { $title = [string]$Process.MainWindowTitle } catch {}
    try { $startedAt = $Process.StartTime.ToString('o') } catch {}

    return [pscustomobject]@{
        pid              = [int]$Process.Id
        visible          = [bool]($handle -ne 0)
        mainWindowHandle = $handle
        windowTitle      = $title
        startedAt        = $startedAt
    }
}

$instances = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
    Get-ExcelInstanceInfo -Process $_
})

if ($Action -eq 'list') {
    [pscustomobject]@{
        action    = 'list'
        count     = $instances.Count
        instances = @($instances)
    } | ConvertTo-Json -Depth 5
    exit 0
}

$targets = switch ($Action) {
    'close-all' { $instances }
    'close-visible' { @($instances | Where-Object { $_.visible }) }
    'close-invisible' { @($instances | Where-Object { -not $_.visible }) }
}

$closed = @()
$failed = @()
foreach ($instance in $targets) {
    try {
        Stop-Process -Id $instance.pid -Force -ErrorAction Stop
        $closed += [int]$instance.pid
    } catch {
        $failed += [pscustomobject]@{
            pid     = [int]$instance.pid
            message = $_.Exception.Message
        }
    }
}

[pscustomobject]@{
    action       = $Action
    requested    = $targets.Count
    closed       = @($closed)
    failed       = @($failed)
} | ConvertTo-Json -Depth 5

if ($failed.Count -gt 0) { exit 1 }
exit 0