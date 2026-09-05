# List-ExcelProcessWindows.ps1
#
# List OS windows and titles associated with an Excel process PID. This works
# without COM and can reveal multiple window handles owned by one process.

[CmdletBinding()]
param(
    [int[]]$ProcessId
)

$ErrorActionPreference = 'Stop'

if (-not ('ExcelProcessWindowsNative' -as [type])) {
    Add-Type @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class ExcelProcessWindowsNative {
    public sealed class WindowInfo {
        public long Hwnd { get; set; }
        public long ParentHwnd { get; set; }
        public int ProcessId { get; set; }
        public bool Visible { get; set; }
        public string ClassName { get; set; }
        public string Title { get; set; }
    }

    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int maxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hwnd, StringBuilder className, int maxCount);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hwnd);

    public static List<WindowInfo> FindWindows(int processId, bool includeChildren) {
        var results = new List<WindowInfo>();
        EnumWindows((hwnd, unused) => {
            AddWindow(hwnd, IntPtr.Zero, processId, results);
            if (includeChildren) {
                EnumChildWindows(hwnd, (child, ignored) => {
                    AddWindow(child, hwnd, processId, results);
                    return true;
                }, IntPtr.Zero);
            }
            return true;
        }, IntPtr.Zero);
        return results;
    }

    private static void AddWindow(IntPtr hwnd, IntPtr parent, int requestedPid, List<WindowInfo> results) {
        uint processId;
        GetWindowThreadProcessId(hwnd, out processId);
        if (processId != requestedPid) return;

        var title = new StringBuilder(1024);
        var className = new StringBuilder(256);
        GetWindowText(hwnd, title, title.Capacity);
        GetClassName(hwnd, className, className.Capacity);
        results.Add(new WindowInfo {
            Hwnd = hwnd.ToInt64(),
            ParentHwnd = parent.ToInt64(),
            ProcessId = (int)processId,
            Visible = IsWindowVisible(hwnd),
            ClassName = className.ToString(),
            Title = title.ToString()
        });
    }
}
"@
}

if (-not $ProcessId) {
    $ProcessId = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
}

if ($ProcessId.Count -eq 0) {
    Write-Output 'No Excel processes found.'
    exit 0
}

foreach ($pidValue in $ProcessId) {
    $process = Get-Process -Id $pidValue -ErrorAction SilentlyContinue
    if ($null -eq $process -or $process.ProcessName -ne 'EXCEL') {
        Write-Output "Skipping PID $pidValue because it is not a running EXCEL process."
        continue
    }

    $windows = @([ExcelProcessWindowsNative]::FindWindows($pidValue, $true) | Where-Object { $_.Title })
    Write-Output "PID $pidValue ($($process.ProcessName)) windows with titles: $($windows.Count)"

    foreach ($window in $windows) {
        $scope = if ($window.ParentHwnd -eq 0) { 'top-level' } else { 'child' }
        Write-Output ("  {0} hwnd={1} parent={2} visible={3} class={4} title={5}" -f $scope, $window.Hwnd, $window.ParentHwnd, $window.Visible, $window.ClassName, $window.Title)
    }
}
