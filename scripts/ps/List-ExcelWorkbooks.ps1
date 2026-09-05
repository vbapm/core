<#
.SYNOPSIS
Lists every open workbook in every running Excel.Application instance.

.DESCRIPTION
[Marshal]::GetActiveObject('Excel.Application') only returns the first instance
registered in the Running Object Table (ROT). This script walks the ROT itself,
picks up every Excel object it can find (workbooks and applications), resolves
each one to its owning Application, dedupes by window handle, then enumerates
Application.Workbooks so workbooks that are not themselves in the ROT (add-ins,
hidden books, PERSONAL.XLSB) are also reported.

.PARAMETER ShowRot
Also print the raw ROT display names, for debugging.
#>
[CmdletBinding()]
param(
    [switch]$ShowRot
)

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public class RotEntry
{
    public string DisplayName;
    public object Object;
}

public static class RotReader
{
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

    [DllImport("user32.dll")]
    private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);

    public static int GetProcessIdFromHwnd(long hwnd)
    {
        int pid;
        GetWindowThreadProcessId(new IntPtr(hwnd), out pid);
        return pid;
    }

    public static List<RotEntry> GetEntries()
    {
        var result = new List<RotEntry>();
        IRunningObjectTable rot;
        if (GetRunningObjectTable(0, out rot) != 0 || rot == null) return result;

        IEnumMoniker enumMoniker;
        rot.EnumRunning(out enumMoniker);
        enumMoniker.Reset();

        var monikers = new IMoniker[1];
        while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
        {
            var entry = new RotEntry();
            IBindCtx ctx;
            CreateBindCtx(0, out ctx);
            try
            {
                monikers[0].GetDisplayName(ctx, null, out entry.DisplayName);
            }
            catch { entry.DisplayName = "<unreadable>"; }
            finally { Marshal.ReleaseComObject(ctx); }

            try
            {
                rot.GetObject(monikers[0], out entry.Object);
            }
            catch { entry.Object = null; }

            result.Add(entry);
        }
        return result;
    }
}
'@

$entries = [RotReader]::GetEntries()

if ($ShowRot) {
    Write-Host "ROT entries ($($entries.Count)):" -ForegroundColor Cyan
    $entries | ForEach-Object { Write-Host "  $($_.DisplayName)" }
    Write-Host
}

# Resolve every ROT object to its Excel Application, dedupe by Hwnd.
$apps = @{}
foreach ($entry in $entries) {
    if ($null -eq $entry.Object) { continue }
    try {
        $app = $entry.Object.Application
        if ($null -eq $app -or $app.Name -ne 'Microsoft Excel') { continue }
        $hwnd = [long]$app.Hwnd
        if (-not $apps.ContainsKey($hwnd)) { $apps[$hwnd] = $app }
    }
    catch {
        # Not an Office object, or no Application property.
    }
}

if ($apps.Count -eq 0) {
    Write-Warning 'No Excel.Application found in the ROT.'
    return
}

foreach ($hwnd in $apps.Keys | Sort-Object) {
    $app = $apps[$hwnd]
    $processId = [RotReader]::GetProcessIdFromHwnd($hwnd)
    foreach ($wb in @($app.Workbooks)) {
        [pscustomobject]@{
            ProcessId = $processId
            Hwnd      = $hwnd
            Visible   = $app.Visible
            Name      = $wb.Name
            Saved     = $wb.Saved
            ReadOnly  = $wb.ReadOnly
            Path      = $wb.FullName
        }
    }
}
