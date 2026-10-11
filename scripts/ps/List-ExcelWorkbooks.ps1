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

if (-not ('RotReaderExcelWorkbooks' -as [type])) {
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public class RotEntryExcelWorkbooks
{
    public string DisplayName;
    public object Object;
}

public static class RotReaderExcelWorkbooks
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

    public static void Release(object o)
    {
        if (o != null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o);
    }

    public static List<RotEntryExcelWorkbooks> GetEntries()
    {
        var result = new List<RotEntryExcelWorkbooks>();
        IRunningObjectTable rot = null;
        IEnumMoniker enumMoniker = null;
        IBindCtx ctx = null;
        try
        {
            int hr = GetRunningObjectTable(0, out rot);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);
            if (rot == null) return result;

            // ComTypes interface methods throw COMException on failing HRESULTs.
            rot.EnumRunning(out enumMoniker);
            if (enumMoniker == null) return result;
            enumMoniker.Reset();

            hr = CreateBindCtx(0, out ctx);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);

            var monikers = new IMoniker[1];
            while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
            {
                var entry = new RotEntryExcelWorkbooks();
                try
                {
                    try { monikers[0].GetDisplayName(ctx, null, out entry.DisplayName); }
                    catch { entry.DisplayName = "<unreadable>"; }

                    try
                    {
                        hr = rot.GetObject(monikers[0], out entry.Object);
                        if (hr != 0) entry.Object = null;
                    }
                    catch { entry.Object = null; }
                }
                finally
                {
                    Release(monikers[0]);
                    monikers[0] = null;
                }
                result.Add(entry);
            }
            return result;
        }
        finally
        {
            Release(ctx);
            Release(enumMoniker);
            Release(rot);
        }
    }
}
'@
}

$entries = [RotReaderExcelWorkbooks]::GetEntries()

if ($ShowRot) {
    Write-Host "ROT entries ($($entries.Count)):" -ForegroundColor Cyan
    $entries | ForEach-Object { Write-Host "  $($_.DisplayName)" }
    Write-Host
}

# Resolve every ROT object to its Excel Application, dedupe by Hwnd.
$apps = @{}
try {
    foreach ($entry in $entries) {
        if ($null -eq $entry.Object) { continue }
        $entryObject = $entry.Object
        $entry.Object = $null
        $app = $null
        $retainEntryObject = $false
        try {
            $entryName = ''
            try { $entryName = [string]$entryObject.Name } catch {}
            if ($entryName -eq 'Microsoft Excel') {
                $app = $entryObject
                $retainEntryObject = $true
            } else {
                $app = $entryObject.Application
            }

            if ($null -eq $app -or [string]$app.Name -ne 'Microsoft Excel') { continue }
            $hwnd = [long]$app.Hwnd
            if (-not $apps.ContainsKey($hwnd)) {
                $apps[$hwnd] = $app
            }

            # The dictionary owns the application reference, including when a
            # duplicate ROT moniker resolves to the same application.
            $app = $null
            if ($retainEntryObject) {
                $entryObject = $null
            }
        }
        catch {
            # Not an Office object, or no Application property.
        }
        finally {
            [RotReaderExcelWorkbooks]::Release($app)
            [RotReaderExcelWorkbooks]::Release($entryObject)
        }
    }

    if ($apps.Count -eq 0) {
        Write-Warning 'No Excel.Application found in the ROT.'
        return
    }

    foreach ($hwnd in $apps.Keys | Sort-Object) {
        $app = $apps[$hwnd]
        $processId = [RotReaderExcelWorkbooks]::GetProcessIdFromHwnd($hwnd)
        $visible = $null
        try { $visible = $app.Visible } catch { }

        $workbooks = $null
        try {
            $workbooks = $app.Workbooks
            $count = $workbooks.Count
            for ($i = 1; $i -le $count; $i++) {
                $wb = $null
                try {
                    $wb = $workbooks.Item($i)
                    [pscustomobject]@{
                        ProcessId = $processId
                        Hwnd      = $hwnd
                        Visible   = $visible
                        Name      = $wb.Name
                        Saved     = $wb.Saved
                        ReadOnly  = $wb.ReadOnly
                        Path      = if ($wb.Path) { $wb.FullName } else { $null }
                    }
                }
                catch {
                    Write-Warning "PID ${processId}: workbook #$i is inaccessible: $($_.Exception.Message)"
                }
                finally {
                    [RotReaderExcelWorkbooks]::Release($wb)
                }
            }
        }
        catch {
            Write-Warning "PID ${processId}: cannot enumerate workbooks: $($_.Exception.Message)"
        }
        finally {
            [RotReaderExcelWorkbooks]::Release($workbooks)
        }
    }
}
finally {
    foreach ($app in $apps.Values) { [RotReaderExcelWorkbooks]::Release($app) }
    $apps.Clear()
}
