# Set-ExcelWindowsVisibility.ps1
#
# Shared PID-based visibility implementation for Excel top-level windows.

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, Position = 0)]
    [Alias('Pid')]
    [int]$ProcessId,

    [Parameter(Mandatory)]
    [bool]$Visible
)

$ErrorActionPreference = 'Stop'

if (-not ('SetExcelWindowsVisibilityNative' -as [type])) {
    Add-Type @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class SetExcelWindowsVisibilityNative {
    public sealed class WindowInfo {
        public long Hwnd { get; set; }
        public string ClassName { get; set; }
        public string Title { get; set; }
    }

    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hwnd, int command);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int maxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hwnd, StringBuilder className, int maxCount);

    public static List<WindowInfo> GetTopLevelWindows(int requestedProcessId) {
        var results = new List<WindowInfo>();
        EnumWindows((hwnd, unused) => {
            uint processId;
            GetWindowThreadProcessId(hwnd, out processId);
            if (processId != requestedProcessId) return true;

            var title = new StringBuilder(1024);
            var className = new StringBuilder(256);
            GetWindowText(hwnd, title, title.Capacity);
            GetClassName(hwnd, className, className.Capacity);
            results.Add(new WindowInfo {
                Hwnd = hwnd.ToInt64(),
                ClassName = className.ToString(),
                Title = title.ToString()
            });
            return true;
        }, IntPtr.Zero);
        return results;
    }

    public static bool SetVisibility(long hwnd, bool visible) {
        return ShowWindow(new IntPtr(hwnd), visible ? 5 : 0);
    }
}
"@
}

$process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
if ($null -eq $process) {
    throw "No process found with PID $ProcessId."
}
if ($process.ProcessName -ne 'EXCEL') {
    throw "PID $ProcessId belongs to $($process.ProcessName), not EXCEL."
}

$windows = @([SetExcelWindowsVisibilityNative]::GetTopLevelWindows($ProcessId))
$action = if ($Visible) { 'Show window' } else { 'Hide window' }
$pastTense = if ($Visible) { 'Shown' } else { 'Hidden' }

if ($windows.Count -eq 0) {
    Write-Output "Excel PID $ProcessId has no top-level windows."
    exit 0
}

$changed = 0
foreach ($window in $windows) {
    $title = if ($window.Title) { $window.Title } else { '<untitled>' }
    $target = "Excel PID $ProcessId hwnd=$($window.Hwnd) class=$($window.ClassName) title=$title"

    if ($PSCmdlet.ShouldProcess($target, $action)) {
        [SetExcelWindowsVisibilityNative]::SetVisibility($window.Hwnd, $Visible) | Out-Null
        Write-Output "$pastTense $target"
        $changed++
    } else {
        Write-Output "Would $($action.ToLowerInvariant()) $target"
    }
}

if ($WhatIfPreference) {
    Write-Output "No windows changed. Excel PID $ProcessId has $($windows.Count) top-level window(s)."
} else {
    Write-Output "$pastTense $changed top-level window(s) for Excel PID $ProcessId."
}
