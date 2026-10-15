[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class VbeWindowStateApi
{
    public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct Rect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int maxCount);

    [DllImport("user32.dll")]
    public static extern int GetClassName(IntPtr hwnd, StringBuilder text, int maxCount);

    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hwnd);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hwnd);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);

    [DllImport("user32.dll")]
    public static extern IntPtr GetWindowLongPtr(IntPtr hwnd, int index);
}
'@

$windows = [System.Collections.Generic.List[object]]::new()
[VbeWindowStateApi]::EnumWindows({
    param($hwnd, $lParam)

    $title = [Text.StringBuilder]::new(512)
    [VbeWindowStateApi]::GetWindowText($hwnd, $title, $title.Capacity) | Out-Null
    if ($title.ToString().Contains("Microsoft Visual Basic")) {
        $className = [Text.StringBuilder]::new(256)
        [VbeWindowStateApi]::GetClassName($hwnd, $className, $className.Capacity) | Out-Null
        $ownerProcessId = [uint32]0
        $ownerThreadId = [VbeWindowStateApi]::GetWindowThreadProcessId($hwnd, [ref]$ownerProcessId)
        $rect = [VbeWindowStateApi+Rect]::new()
        [VbeWindowStateApi]::GetWindowRect($hwnd, [ref]$rect) | Out-Null
        $style = [VbeWindowStateApi]::GetWindowLongPtr($hwnd, -16).ToInt64()
        $exStyle = [VbeWindowStateApi]::GetWindowLongPtr($hwnd, -20).ToInt64()

        $windows.Add([pscustomobject]@{
            Hwnd = ("0x{0:X8}" -f $hwnd.ToInt64())
            OwnerPid = $ownerProcessId
            OwnerThread = $ownerThreadId
            IsWindow = [VbeWindowStateApi]::IsWindow($hwnd)
            Visible = [VbeWindowStateApi]::IsWindowVisible($hwnd)
            Foreground = $hwnd -eq [VbeWindowStateApi]::GetForegroundWindow()
            Class = $className.ToString()
            Style = ("0x{0:X}" -f $style)
            ExStyle = ("0x{0:X}" -f $exStyle)
            Bounds = ("{0},{1} {2}x{3}" -f $rect.Left, $rect.Top, ($rect.Right - $rect.Left), ($rect.Bottom - $rect.Top))
            Title = $title.ToString()
        })
    }

    return $true
}, [IntPtr]::Zero) | Out-Null

Write-Host "Native VBE windows:" -ForegroundColor Cyan
if ($windows.Count -eq 0) {
    Write-Host "  None found."
} else {
    $windows | Format-List
}

try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    Write-Host "Excel COM state:" -ForegroundColor Cyan
    [pscustomobject]@{
        ExcelHwnd = ("0x{0:X8}" -f $excel.Hwnd)
        VBEVisible = $excel.VBE.MainWindow.Visible
        VBECaption = $excel.VBE.MainWindow.Caption
        WorkbookCount = $excel.Workbooks.Count
    } | Format-List
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
} catch {
    Write-Host "Excel COM state unavailable: $($_.Exception.Message)" -ForegroundColor Yellow
}