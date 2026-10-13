[CmdletBinding()]
param(
    [ValidateSet('list', 'close-all', 'close-visible', 'close-invisible')]
    [string]$Action = 'list',
    [switch]$Force
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

function Get-ExcelApplications {
    $rotClass = 'HostRotHelper'
    if (-not ($rotClass -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class HostRotHelper {
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable rot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(uint reserved, out IBindCtx ctx);

    public static List<object> GetExcelObjects() {
        var result = new List<object>();
        IRunningObjectTable rot = null;
        IBindCtx ctx = null;
        IEnumMoniker monikers = null;
        if (GetRunningObjectTable(0, out rot) != 0 || rot == null) return result;
        if (CreateBindCtx(0, out ctx) != 0 || ctx == null) return result;

        try {
            rot.EnumRunning(out monikers);
            if (monikers == null) return result;
            monikers.Reset();
            var fetched = new IMoniker[1];
            while (monikers.Next(1, fetched, IntPtr.Zero) == 0) {
                object value = null;
                try {
                    if (rot.GetObject(fetched[0], out value) == 0 && value != null) {
                        result.Add(value);
                        value = null;
                    }
                } catch {}
                finally {
                    if (value != null) { try { Marshal.ReleaseComObject(value); } catch {} }
                    if (fetched[0] != null) { try { Marshal.ReleaseComObject(fetched[0]); } catch {} }
                }
            }
            return result;
        }
        finally {
            if (monikers != null) Marshal.ReleaseComObject(monikers);
            if (ctx != null) Marshal.ReleaseComObject(ctx);
            if (rot != null) Marshal.ReleaseComObject(rot);
        }
    }

    public static int GetProcessId(long hwnd) {
        uint pid;
        GetWindowThreadProcessId(new IntPtr(hwnd), out pid);
        return (int)pid;
    }

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
}
"@
    }

    $objects = @([HostRotHelper]::GetExcelObjects())
    $apps = @{}
    foreach ($candidate in $objects) {
        $app = $null
        $ownsApp = $false
        try {
            try {
                if ([string]$candidate.Name -eq 'Microsoft Excel') {
                    $app = $candidate
                    $ownsApp = $true
                }
            } catch {}
            if ($null -eq $app) { $app = $candidate.Application }
            if ($null -eq $app) { continue }
            if ([string]$app.Name -ne 'Microsoft Excel') { continue }

            $key = [string]$app.Hwnd
            if (-not $apps.ContainsKey($key)) { $apps[$key] = $app; $app = $null }
        } catch {}
        finally {
            if ($app -and -not $ownsApp) { try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null } catch {} }
            if (-not $ownsApp) { try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($candidate) | Out-Null } catch {} }
        }
    }

    return @($apps.Values)
}

function Get-VisibleUnsavedWorkbooks {
    param([Parameter(Mandatory)][object[]]$VisibleInstances)

    $visiblePids = @{}
    foreach ($instance in $VisibleInstances) { $visiblePids[[int]$instance.pid] = $true }
    $inspectedPids = @{}
    $unsaved = @()

    foreach ($app in @(Get-ExcelApplications)) {
        try {
            $pid = [HostRotHelper]::GetProcessId([int64]$app.Hwnd)
            if (-not $visiblePids.ContainsKey($pid)) { continue }
            $inspectedPids[$pid] = $true
            foreach ($workbook in @($app.Workbooks)) {
                try {
                    if (-not [bool]$workbook.Saved) {
                        $unsaved += [pscustomobject]@{
                            pid      = $pid
                            name     = [string]$workbook.Name
                            fullName = [string]$workbook.FullName
                        }
                    }
                } catch {
                    $unsaved += [pscustomobject]@{
                        pid      = $pid
                        name     = '(unreadable workbook)'
                        fullName = ''
                    }
                }
            }
        } catch {}
        finally {
            try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null } catch {}
        }
    }

    $uninspected = @($visiblePids.Keys | Where-Object { -not $inspectedPids.ContainsKey([int]$_) })
    return [pscustomobject]@{ unsaved = @($unsaved); uninspected = @($uninspected) }
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

if (-not $Force -and $targets.Count -gt 0 -and @($targets | Where-Object { $_.visible }).Count -gt 0) {
    $check = Get-VisibleUnsavedWorkbooks -VisibleInstances @($targets | Where-Object { $_.visible })
    if ($check.unsaved.Count -gt 0 -or $check.uninspected.Count -gt 0) {
        [pscustomobject]@{
            action      = $Action
            error       = 'Visible Excel instances have unsaved workbooks or could not be inspected. Re-run with --force to terminate them.'
            unsaved     = @($check.unsaved)
            uninspected = @($check.uninspected)
        } | ConvertTo-Json -Depth 5
        exit 2
    }
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