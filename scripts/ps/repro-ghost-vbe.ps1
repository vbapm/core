<#
.SYNOPSIS
    Reproduces the ghost VBE project bug.

.DESCRIPTION
    Automates the full repro:
      1. Close all Excel instances
      2. Open a fresh visible Excel instance (with blank workbook)
      3. Open the VBE window (Alt+F11)
      4. Open the project workbook via vba open
      5. Close the workbook via COM
      6. Check for ghost VBE projects

.PARAMETER ProjectDir
    Path to the vbapm project to test with.
    Default: demo/simple-demo-thisworkbook5

.PARAMETER SkipVBE
    Don't open the VBE window (test without VBE visible).

.EXAMPLE
    .\scripts\ps\repro-ghost-vbe.ps1

.EXAMPLE
    .\scripts\ps\repro-ghost-vbe.ps1 -SkipVBE
#>

param(
    [string]$ProjectDir = "demo/simple-demo-thisworkbook5",
    [switch]$SkipVBE
)

# ---------------------------------------------------------------
# PowerShell Core -> re-invoke with Windows PowerShell (COM+Win32)
# ---------------------------------------------------------------
if ($PSVersionTable.PSEdition -eq 'Core') {
    $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
    if ($SkipVBE)    { $psArgs += '-SkipVBE' }
    $psArgs += '-ProjectDir'; $psArgs += $ProjectDir
    & powershell.exe $psArgs
    exit $LASTEXITCODE
}

$ErrorActionPreference = 'Stop'
$repoRoot = $PSScriptRoot + "/../.."
$scripts  = Join-Path $repoRoot "scripts/ps"

# Win32 helpers
$Win32 = Add-Type -MemberDefinition @'
[DllImport("user32.dll", SetLastError=true)]
public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
[DllImport("user32.dll")]
public static extern bool IsWindow(IntPtr hWnd);
[DllImport("user32.dll")]
public static extern bool SetForegroundWindow(IntPtr hWnd);
'@ -Name "Win32VBE" -Namespace "Util" -PassThru

$VBE_CLASS = "wndclass_desked_gsk"

function Find-VBEWindow {
    # Try by class name first
    $hwnd = $Win32::FindWindow($VBE_CLASS, $null)
    if ($hwnd -ne [IntPtr]::Zero) { return $hwnd }

    # Fallback: enumerate top-level windows looking for "Microsoft Visual Basic"
    $result = [IntPtr]::Zero
    $filter = "Microsoft Visual Basic"
    Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class WindowFinder {
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    public static IntPtr FindByTitleContains(string partialTitle) {
        IntPtr found = IntPtr.Zero;
        EnumWindows((hWnd, lParam) => {
            StringBuilder sb = new StringBuilder(256);
            GetWindowText(hWnd, sb, 256);
            if (sb.ToString().Contains(partialTitle)) {
                found = hWnd;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
'@ -ReferencedAssemblies "System.Runtime.InteropServices"
    return [WindowFinder]::FindByTitleContains($filter)
}

function Open-VBE {
    Write-Host "Opening VBE window..." -ForegroundColor Cyan

    # Try COM first (needs Trust Center setting)
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        $excel.VBE.MainWindow.Visible = $true
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        Start-Sleep -Seconds 2
    } catch {
        # Fallback: force Excel to foreground then SendKeys
        Write-Host "  COM VBE access failed - using SendKeys..." -ForegroundColor Yellow
        try {
            $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            $hwndExcel = $excel.Hwnd
            if ($hwndExcel -ne [IntPtr]::Zero) {
                $Win32::SetForegroundWindow($hwndExcel) | Out-Null
                Start-Sleep -Milliseconds 300
                [System.Windows.Forms.SendKeys]::SendWait("%{F11}")
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        } catch {
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.SendKeys]::SendWait("%{F11}")
        }
        Start-Sleep -Seconds 2
    }

    $hwnd = Find-VBEWindow
    if ($hwnd -ne [IntPtr]::Zero) {
        Write-Host "  VBE window open: HWND 0x$($hwnd.ToString('X8'))" -ForegroundColor Green
    } else {
        Write-Host "  VBE window NOT found (tried class name + title search)." -ForegroundColor Red
    }
    return ($hwnd -ne [IntPtr]::Zero)
}

function Step([string]$label) {
    Write-Host ""
    Write-Host "===== $label =====" -ForegroundColor Cyan
}

function Run-Script([string]$name, [string]$extraArgs = "") {
    $path = Join-Path $scripts $name
    Invoke-Expression "& '$path' $extraArgs"
    if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) { throw "$name failed (exit $LASTEXITCODE)" }
}

function Run-Cmd([string]$cmd) {
    Write-Host "> $cmd" -ForegroundColor DarkGray
    Invoke-Expression $cmd
    if ($LASTEXITCODE -ne 0) { throw "Command failed: $cmd" }
}

Push-Location $repoRoot
try {
    # ---------------------------------------------------------------
    Step "1. Close all Excel instances"
    # ---------------------------------------------------------------
    Run-Script "close-excel.ps1" "-SaveAll"
    Run-Script "close-hidden-excel-instances.ps1" "-Force"

    $procs = @(Get-Process excel -ErrorAction SilentlyContinue)
    if ($procs.Count -gt 0) {
        Write-Host "WARNING: $($procs.Count) Excel process(es) still running. Force-killing..." -ForegroundColor Yellow
        $procs | Stop-Process -Force
        Start-Sleep -Seconds 2
    }
    Write-Host "All Excel instances closed." -ForegroundColor Green

    # ---------------------------------------------------------------
    Step "2. Open new visible Excel instance (blank workbook)"
    # ---------------------------------------------------------------
    Run-Script "open-new-excel.ps1"
    Start-Sleep -Seconds 3
    Write-Host "Excel ready." -ForegroundColor Green

    # ---------------------------------------------------------------
    Step "3. Open VBE window"
    # ---------------------------------------------------------------
    if ($SkipVBE) {
        Write-Host "Skipped (-SkipVBE)" -ForegroundColor DarkGray
    } else {
        $vbeOpen = Open-VBE
        if (-not $vbeOpen) {
            Write-Host "Proceeding anyway - VBE may still open when workbook is loaded."
        }
    }

    # ---------------------------------------------------------------
    Step "4. Run vba open"
    # ---------------------------------------------------------------
    Set-Location (Join-Path $repoRoot $ProjectDir)
    Run-Cmd "vba open"
    Start-Sleep -Seconds 2
    Write-Host "Workbook opened via vba open." -ForegroundColor Green

    # Re-verify VBE is still open after opening the workbook
    if (-not $SkipVBE) {
        $vbeHwnd = Find-VBEWindow
        if ($vbeHwnd -ne [IntPtr]::Zero) {
            Write-Host "VBE window still open: HWND 0x$($vbeHwnd.ToString('X8'))" -ForegroundColor Green
        } else {
            Write-Host "WARNING: VBE window not found (may be behind Excel window)." -ForegroundColor Yellow
        }
    }

    Set-Location $repoRoot

    # ---------------------------------------------------------------
    Step "5. Close the workbook via COM"
    # ---------------------------------------------------------------
    $wbName = (Split-Path $ProjectDir -Leaf) + ".xlsm"
    Run-Script "close-workbook.ps1" "-Name $wbName -Force"
    Write-Host "Workbook closed via COM." -ForegroundColor Green

    # ---------------------------------------------------------------
    Step "6. Check for ghost VBE projects"
    # ---------------------------------------------------------------
    Write-Host ""
    Run-Script "detect-ghost-vbe.ps1"

    # ---------------------------------------------------------------
    Write-Host ""
    Write-Host "===== Repro complete =====" -ForegroundColor Cyan
    Write-Host "If a GHOST project appeared above, the bug is reproduced."
    Write-Host "If not, the conditions that trigger it were not met this run."
}
finally {
    Pop-Location
}
