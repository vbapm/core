<#
.SYNOPSIS
    CI-friendly ghost VBE bug repro — creates a fresh project on the fly.

.DESCRIPTION
    Creates everything from scratch:
      1. Close all Excel instances
      2. Open a fresh visible Excel instance
      3. Open the VBE window
      4. Create a new vbapm project via vba init --target xlsm
      5. Add a source module and build
      6. Open the built file via vba open
      7. Close the workbook via COM
      8. Check for ghost VBE projects

.PARAMETER WorkDir
    Directory where the temp project will be created.
    Default: demo/ghost-demo

.EXAMPLE
    .\scripts\ps\repro-ghost-vbe-ci.ps1
#>

param(
    [string]$WorkDir = "ghost-demo"
)

if ($PSVersionTable.PSEdition -eq 'Core') {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -WorkDir $WorkDir
    exit $LASTEXITCODE
}

$ErrorActionPreference = 'Stop'
$repoRoot = $PSScriptRoot + "/../.."
$scripts  = Join-Path $repoRoot "scripts/ps"

$Win32 = Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);' -Name "W32" -Namespace "U" -PassThru

function Find-VBEWindow {
    Add-Type @'
using System;using System.Runtime.InteropServices;using System.Text;
public static class Wf{[DllImport("user32.dll")]static extern bool EnumWindows(E cb,IntPtr l);[DllImport("user32.dll")]static extern int GetWindowText(IntPtr h,StringBuilder t,int m);delegate bool E(IntPtr h,IntPtr l);public static IntPtr ByTitle(string p){IntPtr f=IntPtr.Zero;EnumWindows((h,_)=>{var s=new StringBuilder(256);GetWindowText(h,s,256);if(s.ToString().Contains(p)){f=h;return false;}return true;},IntPtr.Zero);return f;}}
'@
    return [Wf]::ByTitle("Microsoft Visual Basic")
}

function Open-VBE {
    Write-Host "Opening VBE window..." -ForegroundColor Cyan
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        $excel.VBE.MainWindow.Visible = $true
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        Start-Sleep -Seconds 2
    } catch {
        Write-Host "  COM VBE failed - using SendKeys..." -ForegroundColor Yellow
        try {
            $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            $h = $excel.Hwnd; if ($h) { $Win32::SetForegroundWindow($h) | Out-Null; Start-Sleep -Milliseconds 300 }
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.SendKeys]::SendWait("%{F11}")
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        } catch { Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait("%{F11}") }
        Start-Sleep -Seconds 2
    }
    $hwnd = Find-VBEWindow
    if ($hwnd -ne [IntPtr]::Zero) { Write-Host "  VBE window open: HWND 0x$($hwnd.ToString('X8'))" -ForegroundColor Green }
    else { Write-Host "  VBE window NOT found." -ForegroundColor Red }
    return ($hwnd -ne [IntPtr]::Zero)
}

function Step([string]$l) { Write-Host ""; Write-Host "===== $l =====" -ForegroundColor Cyan }

function Run-Cmd([string]$cmd, [string]$cwd) {
    Write-Host "> $cmd" -ForegroundColor DarkGray
    if ($cwd) { Push-Location $cwd }
    try { Invoke-Expression $cmd } finally { if ($cwd) { Pop-Location } }
    if ($LASTEXITCODE -ne 0) { throw "Command failed (exit $LASTEXITCODE): $cmd" }
}

Push-Location $repoRoot
try {
    Step "1. Close all Excel instances"
    & (Join-Path $scripts "close-excel.ps1") -SaveAll
    & (Join-Path $scripts "close-hidden-excel-instances.ps1") -Force
    Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2
    Write-Host "All Excel instances closed." -ForegroundColor Green

    Step "2. Open new visible Excel instance"
    & (Join-Path $scripts "open-new-excel.ps1")
    Start-Sleep -Seconds 3
    Write-Host "Excel ready." -ForegroundColor Green

    Step "3. Open VBE window"
    $vbeOpen = Open-VBE

    Step "4. Create fresh vbapm project"
    $projectDir = Join-Path $repoRoot $WorkDir
    if (Test-Path $projectDir) { Remove-Item -Recurse -Force $projectDir }
    New-Item -ItemType Directory -Path $projectDir -Force | Out-Null
    Run-Cmd "vba init --target xlsm" $projectDir
    Run-Cmd "vba add Module1" $projectDir

    # vba add generates absolute paths and src-subfolders globs —
    # rewrite manifest with minimal relative paths that actually build
    @"
[project]
name = "ghost-demo"
version = "0.0.0"
authors = []
license = "UNLICENSED"
target = "xlsm"

[src]
Module1 = "src/Module1.bas"
"@ | Set-Content (Join-Path $projectDir "vbaproject.toml")
    Write-Host "Project initialized." -ForegroundColor Green

    Step "5. Build the project"
    Run-Cmd "vba build" $projectDir
    Write-Host "Project built." -ForegroundColor Green

    Step "6. Open the built file"
    # Open via COM (Workbooks.Open) instead of `vba open`.
    # `vba open` shells out to `Start -Wait` which hangs in headless CI
    # because the .xlsm file association never returns while Excel stays open.
    # The ghost bug triggers on any open+close path, so COM open is equivalent.
    $projLeaf = Split-Path $WorkDir -Leaf
    $builtFile = Join-Path $projectDir "build/$projLeaf.xlsm"
    $microsoftExcel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $microsoftExcel.Workbooks.Open($builtFile) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($microsoftExcel) | Out-Null
    Start-Sleep -Seconds 2
    Write-Host "Workbook opened via COM (Workbooks.Open)." -ForegroundColor Green
    if ($vbeOpen) {
        $vh = Find-VBEWindow
        if ($vh) { Write-Host "VBE still open: HWND 0x$($vh.ToString('X8'))" -ForegroundColor Green }
        else { Write-Host "VBE not found after open." -ForegroundColor Yellow }
    }

    Step "7. Close the workbook via COM"
    $wbName = (Split-Path $WorkDir -Leaf) + ".xlsm"
    & (Join-Path $scripts "close-workbook.ps1") -Name $wbName -Force
    Write-Host "Workbook closed via COM." -ForegroundColor Green

    Step "8. Check for ghost VBE projects"
    Write-Host ""
    & (Join-Path $scripts "detect-ghost-vbe.ps1")
    Write-Host ""; Write-Host "===== Repro complete =====" -ForegroundColor Cyan
} finally { Pop-Location }
