# Get-WindowChildren.ps1
#
# Read-only inspection of the child windows owned by a parent window.
#
# Examples:
#   .\Get-WindowChildren.ps1 -ParentHandle 3804456
#   .\Get-WindowChildren.ps1 -ParentHandle 3804456 -VisibleOnly -TextOnly
#   .\Get-WindowChildren.ps1 -ParentHandle 3804456 -ClassPattern 'button|static'

[CmdletBinding()]
param(
    [Parameter(Mandatory)][long]$ParentHandle,
    [switch]$VisibleOnly,
    [switch]$TextOnly,
    [string]$ClassPattern,
    [string]$TextPattern
)

$ErrorActionPreference = 'Stop'

if (-not ('WindowChildrenNative' -as [type])) {
    Add-Type @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public static class WindowChildrenNative {
    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr parameter);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr window, StringBuilder className, int capacity);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr window);

    public delegate bool EnumProc(IntPtr window, IntPtr parameter);

    public static object[] GetChildren(long parentHandle) {
        var rows = new List<object>();
        EnumChildWindows((IntPtr)parentHandle, (window, parameter) => {
            var text = new StringBuilder(1024);
            var className = new StringBuilder(256);
            GetWindowText(window, text, text.Capacity);
            GetClassName(window, className, className.Capacity);
            rows.Add(new {
                Handle = window.ToInt64(),
                Visible = IsWindowVisible(window),
                Class = className.ToString(),
                Text = text.ToString()
            });
            return true;
        }, IntPtr.Zero);
        return rows.ToArray();
    }
}
'@
}

$windows = [WindowChildrenNative]::GetChildren($ParentHandle)
foreach ($window in $windows) {
    if ($VisibleOnly -and -not $window.Visible) { continue }
    if ($TextOnly -and [string]::IsNullOrEmpty($window.Text)) { continue }
    if ($ClassPattern -and $window.Class -notmatch $ClassPattern) { continue }
    if ($TextPattern -and $window.Text -notmatch $TextPattern) { continue }

    [pscustomobject]@{
        Handle  = $window.Handle
        Visible = $window.Visible
        Class   = $window.Class
        Text    = $window.Text
    }
}
