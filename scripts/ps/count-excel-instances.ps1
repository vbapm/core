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
# COM workbooks snapshot (from the first ROT-registered instance)
# ---------------------------------------------------------------
Write-Host "===== COM Snapshot (first ROT instance) =====" -ForegroundColor Cyan
try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")

    # Resolve which PID owns the ROT instance via its window handle
    $rotPid = "?"
    try {
        $sig = '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);'
        $Win32 = Add-Type -MemberDefinition $sig -Name "Win32User" -Namespace "Util" -PassThru
        $pidOut = 0
        $Win32::GetWindowThreadProcessId($excel.Hwnd, [ref]$pidOut) | Out-Null
        if ($pidOut) { $rotPid = $pidOut }
    } catch { }

    Write-Host "ROT instance PID: $rotPid"
    Write-Host "Workbooks open in this instance: $($excel.Workbooks.Count)"
    foreach ($wb in $excel.Workbooks) {
        $saved = if ($wb.Saved) { "saved" } else { "MODIFIED" }
        Write-Host "  $($wb.Name)  [$saved]  $($wb.FullName)"
    }
    Write-Host ""
    Write-Host "Add-ins loaded:"
    foreach ($ai in $excel.AddIns) {
        if ($ai.Installed) {
            Write-Host "  $($ai.Name)  (installed)"
        }
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

# ---------------------------------------------------------------
# Summary
# ---------------------------------------------------------------
Write-Host ""
Write-Host "===== Summary =====" -ForegroundColor Cyan
Write-Host "Total Excel processes: $($procs.Count)"
Write-Host "Note: If > 1 instance is running, only the FIRST one's workbooks"
Write-Host "      are visible via COM (GetActiveObject returns the first ROT entry)."

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
