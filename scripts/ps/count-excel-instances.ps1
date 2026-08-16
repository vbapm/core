<#
.SYNOPSIS
    Counts running Excel instances and shows what each one has open.
#>

param(
    [switch]$Raw
)

# ---------------------------------------------------------------
# PowerShell Core → re-invoke with Windows PowerShell for COM
# ---------------------------------------------------------------
if ($PSVersionTable.PSEdition -eq 'Core') {
    $winPS = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($winPS) {
        $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
        if ($Raw) { $psArgs += '-Raw' }
        & powershell.exe $psArgs
        exit $LASTEXITCODE
    }
}

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------
# Get process info
# ---------------------------------------------------------------
$procs = @(Get-Process excel -ErrorAction SilentlyContinue)

Write-Host "===== Excel Processes =====" -ForegroundColor Cyan
Write-Host "Process count : $($procs.Count)"
Write-Host ""

$entries = @()

foreach ($proc in $procs) {
    $procId    = $proc.Id
    $startTime = $proc.StartTime
    $memoryMB  = [math]::Round($proc.WorkingSet64 / 1MB, 1)
    $mainTitle = $proc.MainWindowTitle
    $visible   = if ($mainTitle) { "visible" } else { "hidden" }

    Write-Host "PID $procId  |  $visible  |  Started: $startTime  |  Mem: ${memoryMB}MB" -ForegroundColor Yellow
    if ($mainTitle) {
        Write-Host "         Main window: ""$mainTitle"""
    }

    # Try to attach via COM and list workbooks in this instance
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        # Note: GetActiveObject always returns the FIRST registered instance.
        # There's no direct way to map a PID to a specific COM instance.
        # We report the workbooks from the first instance but flag the ambiguity.
    } catch { }

    $entries += [PSCustomObject]@{
        PID        = $procId
        StartTime  = $startTime
        Visible    = $visible
        MemoryMB   = $memoryMB
        MainTitle  = $mainTitle
    }

    Write-Host ""
}

# ---------------------------------------------------------------
# COM workbooks snapshot — ROT-registered instance
# ---------------------------------------------------------------
# Only one Excel instance registers in the ROT.  GetActiveObject
# always returns that one.  Other instances exist as processes but
# are invisible to COM (no API to reach them by PID alone).
#
# Sidenote: AccessibleObjectFromWindow
# was tested with OBJID_NATIVEOM (E_FAIL for XLMAIN) and OBJID_WINDOW
# (returns MSAA wrapper, not Excel.Application) — neither exposes the
# native OM for other instances.

Write-Host "===== COM Snapshot (ROT-registered instance) =====" -ForegroundColor Cyan

$snapCount = 0
$rotPid    = "?"

try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")

    # Resolve PID from HWND
    try {
        $Win32 = Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);' -Name "Win32User" -Namespace "Util" -PassThru
        $pidOut = 0
        $Win32::GetWindowThreadProcessId($excel.Hwnd, [ref]$pidOut) | Out-Null
        if ($pidOut) { $rotPid = $pidOut }
    } catch { }

    $snapCount++
    Write-Host "PID $rotPid" -ForegroundColor Yellow
    Write-Host "Workbooks: $($excel.Workbooks.Count)"
    foreach ($wb in $excel.Workbooks) {
        $saved = if ($wb.Saved) { "saved" } else { "MODIFIED" }
        Write-Host "  $($wb.Name)  [$saved]  $($wb.FullName)"
    }

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
} catch {
    $msg = $_.Exception.Message
    if ($msg -match "MK_E_UNAVAILABLE") {
        Write-Host "(no ROT entry - process may be a zombie mid-shutdown)"
    } else {
        Write-Host "(could not attach to any COM instance: $msg)"
    }
}

# List other instances that are NOT COM-reachable
$otherProcs = @($procs | Where-Object { $_.Id -ne $rotPid })
if ($otherProcs.Count -gt 0) {
    Write-Host ""
    Write-Host "Other instance(s) - NOT COM-reachable:" -ForegroundColor DarkGray
    foreach ($p in $otherProcs) {
        $msg = "  PID " + $p.Id + "  |  "
        if ($p.MainWindowTitle) {
            $msg += "visible  |  " + $p.MainWindowTitle
        } else {
            $msg += "hidden"
        }
        Write-Host $msg -ForegroundColor DarkGray
    }
    Write-Host "  (vbapm can only target the ROT instance above)" -ForegroundColor DarkGray
}

# ---------------------------------------------------------------
# Summary
# ---------------------------------------------------------------
Write-Host ""
Write-Host "===== Summary =====" -ForegroundColor Cyan
Write-Host "Total Excel processes: $($procs.Count)"
Write-Host "COM-reachable (ROT): $snapCount  (PID $rotPid)"

if ($procs.Count -gt 1) {
    Write-Host ""
    Write-Host "WARNING: $($procs.Count) Excel instances detected!" -ForegroundColor Red
    Write-Host "Multiple instances can cause ghost VBE projects because vbapm"
    Write-Host "operations may target a different instance than the visible one."
    Write-Host ""
    Write-Host "To clean up: close all Excel windows, then check Task Manager"
    Write-Host "for leftover EXCEL.EXE processes and end them manually."
}

if ($Raw) {
    ConvertTo-Json $entries -Depth 2
}
