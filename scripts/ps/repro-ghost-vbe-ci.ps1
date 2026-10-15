<#
.SYNOPSIS
    Demonstrates that opening the VBE before a workbook and then closing that
    workbook can leave its VBA project as a ghost in the VBE.

.DESCRIPTION
        This reproduces an Excel/VBE lifecycle bug, not a vbapm build failure.
        The VBE is opened first, a workbook with the explicit VBA project codename
        "GhostDemoProject" is created and opened, and the workbook is then closed
        while Excel remains running. By default the standalone COM helper closes
        it; -UseCliClose uses the installed vbapm close command and its VBA macro.
        The final detector verifies whether the project remains in the VBE even
        though its workbook is closed.

        Creates everything from scratch:
            1. Close all Excel instances
            2. Create a new vbapm project via vba init --target xlsm
            3. Add a source module and build
            4. Open a fresh visible Excel instance
            5. Open the VBE window
            6. Open the built file via COM after the VBE is loaded
            7. Close the workbook through the selected path
            8. Check for the GhostDemoProject ghost

.PARAMETER WorkDir
    Directory where the temp project will be created.
    Default: ghost-demo

.PARAMETER HideVBEBeforeClose
    Hide the VBE window immediately before closing the workbook.

.PARAMETER CloseVBEBeforeClose
    Close the VBE window immediately before closing the workbook.

.PARAMETER UseCliClose
    Close the workbook through the installed vba close command.

.PARAMETER CloseAndReopenVBEAfterClose
    After closing the workbook, close the VBE through a verified Win32 window
    message, reopen it through VBIDE COM, and check whether the ghost clears.

.PARAMETER RefreshVBEAfterClose
    After closing the workbook, toggle the VBE visibility off and on through
    VBIDE COM, then check whether the ghost clears.

.PARAMETER HoldAfterDetection
    Keep COM references to the Excel VBE and its projects alive after the final
    detector scan so the state can be inspected manually.

.EXAMPLE
    .\scripts\ps\repro-ghost-vbe-ci.ps1

.EXAMPLE
    .\scripts\ps\repro-ghost-vbe-ci.ps1 -UseCliClose -WorkDir ghost-demo-vba
#>

param(
    [string]$WorkDir = "ghost-demo",
    [switch]$HideVBEBeforeClose,
    [switch]$CloseVBEBeforeClose,
    [switch]$UseCliClose,
    [switch]$CloseAndReopenVBEAfterClose,
    [switch]$RefreshVBEAfterClose,
    [switch]$HoldAfterDetection
)

if ($PSVersionTable.PSEdition -eq 'Core') {
    $psArgs = @(
        '-NoProfile'
        '-ExecutionPolicy'
        'Bypass'
        '-File'
        $PSCommandPath
        '-WorkDir'
        $WorkDir
    )
    if ($HideVBEBeforeClose) { $psArgs += '-HideVBEBeforeClose' }
    if ($CloseVBEBeforeClose) { $psArgs += '-CloseVBEBeforeClose' }
    if ($UseCliClose) { $psArgs += '-UseCliClose' }
    if ($CloseAndReopenVBEAfterClose) { $psArgs += '-CloseAndReopenVBEAfterClose' }
    if ($RefreshVBEAfterClose) { $psArgs += '-RefreshVBEAfterClose' }
    if ($HoldAfterDetection) { $psArgs += '-HoldAfterDetection' }
    & powershell.exe @psArgs
    exit $LASTEXITCODE
}

$ErrorActionPreference = 'Stop'
$repoRoot = $PSScriptRoot + "/../.."
$scripts  = Join-Path $repoRoot "scripts/ps"

$Win32 = ([System.Management.Automation.PSTypeName]'U.W32').Type
if (-not $Win32) {
    $Win32 = Add-Type -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern bool SetActiveWindow(IntPtr hWnd);
[DllImport("user32.dll")] public static extern IntPtr SetFocus(IntPtr hWnd);
[DllImport("user32.dll")] public static extern IntPtr GetFocus();
[DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr hWnd, uint flags);
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
[DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr insertAfter, int x, int y, int width, int height, uint flags);
[DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
[DllImport("user32.dll")] public static extern bool AttachThreadInput(uint sourceThreadId, uint targetThreadId, bool attach);
[DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
'@ -Name "W32" -Namespace "U" -PassThru
}

if (-not ([System.Management.Automation.PSTypeName]'W32Input').Type) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;

public static class W32Input {
    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT {
        public uint type;
        public INPUTUNION input;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct INPUTUNION {
        [FieldOffset(0)] public MOUSEINPUT mouse;
        [FieldOffset(0)] public KEYBDINPUT keyboard;
        [FieldOffset(0)] public HARDWAREINPUT hardware;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT {
        public int x;
        public int y;
        public uint mouseData;
        public uint flags;
        public uint time;
        public UIntPtr extraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT {
        public ushort virtualKey;
        public ushort scanCode;
        public uint flags;
        public uint time;
        public UIntPtr extraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct HARDWAREINPUT {
        public uint message;
        public ushort parameterLow;
        public ushort parameterHigh;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint count, INPUT[] inputs, int size);

    public static bool SendAltQ() {
        const uint INPUT_KEYBOARD = 1;
        const uint KEYEVENTF_KEYUP = 2;
        const ushort VK_MENU = 0x12;
        const ushort VK_Q = 0x51;
        var inputs = new[] {
            new INPUT { type = INPUT_KEYBOARD, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = VK_MENU } } },
            new INPUT { type = INPUT_KEYBOARD, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = VK_Q } } },
            new INPUT { type = INPUT_KEYBOARD, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = VK_Q, flags = KEYEVENTF_KEYUP } } },
            new INPUT { type = INPUT_KEYBOARD, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = VK_MENU, flags = KEYEVENTF_KEYUP } } }
        };
        return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT))) == inputs.Length;
    }

    public static bool SendAltF4() {
        const uint INPUT_KEYBOARD = 1;
        const uint KEYEVENTF_KEYUP = 2;
        const ushort VK_MENU = 0x12;
        const ushort VK_F4 = 0x73;
        var inputs = new[] {
            new INPUT { type = INPUT_KEYBOARD, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = VK_MENU } } },
            new INPUT { type = INPUT_KEYBOARD, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = VK_F4 } } },
            new INPUT { type = INPUT_KEYBOARD, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = VK_F4, flags = KEYEVENTF_KEYUP } } },
            new INPUT { type = INPUT_KEYBOARD, input = new INPUTUNION { keyboard = new KEYBDINPUT { virtualKey = VK_MENU, flags = KEYEVENTF_KEYUP } } }
        };
        return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT))) == inputs.Length;
    }
}
'@
}

$script:ExcelProcessId = [uint32]0

function Find-VBEWindow {
    param([uint32]$ProcessId = $script:ExcelProcessId)

    if (-not $ProcessId) { throw "The target Excel process ID is not set." }
    if (-not ([System.Management.Automation.PSTypeName]'Wf').Type) {
        Add-Type @'
using System;using System.Runtime.InteropServices;using System.Text;
public static class Wf{[DllImport("user32.dll")]static extern bool EnumWindows(E cb,IntPtr l);[DllImport("user32.dll")]static extern int GetWindowText(IntPtr h,StringBuilder t,int m);[DllImport("user32.dll")]static extern bool IsWindowVisible(IntPtr h);[DllImport("user32.dll")]static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);delegate bool E(IntPtr h,IntPtr l);public static IntPtr ByTitle(string title,uint processId){IntPtr f=IntPtr.Zero;EnumWindows((h,_)=>{var s=new StringBuilder(256);uint ownerProcessId=0;GetWindowText(h,s,256);GetWindowThreadProcessId(h,out ownerProcessId);if(IsWindowVisible(h)&&ownerProcessId==processId&&s.ToString().Contains(title)){f=h;return false;}return true;},IntPtr.Zero);return f;}}
'@
    }
    return [Wf]::ByTitle("Microsoft Visual Basic", $ProcessId)
}

function Close-VBEWindow {
    $hwnd = Find-VBEWindow
    if ($hwnd -eq [IntPtr]::Zero) { throw "VBE window was not found." }

    if (-not $Win32::SetForegroundWindow($hwnd)) {
        throw "SetForegroundWindow failed for VBE HWND 0x$($hwnd.ToString('X8'))."
    }
    Start-Sleep -Milliseconds 300
    $foreground = $Win32::GetForegroundWindow()
    if ($foreground -ne $hwnd) {
        throw "VBE was not the foreground window. Expected 0x$($hwnd.ToString('X8')), got 0x$($foreground.ToString('X8'))."
    }

    [void]$Win32::SendMessage($hwnd, 0x0112, [IntPtr]0xF060, [IntPtr]::Zero) # WM_SYSCOMMAND / SC_CLOSE
    for ($i = 0; $i -lt 20; $i++) {
        Start-Sleep -Milliseconds 250
        if ((Find-VBEWindow) -eq [IntPtr]::Zero) {
            Write-Host "  VBE closed through verified SC_CLOSE." -ForegroundColor Yellow
            return
        }
    }
    throw "VBE window did not close after SC_CLOSE."
}

function Close-VBEWithAltF4 {
    $hwnd = Find-VBEWindow
    if ($hwnd -eq [IntPtr]::Zero) { throw "VBE window was not found." }

    $targetProcessId = [uint32]0
    $targetThreadId = $Win32::GetWindowThreadProcessId($hwnd, [ref]$targetProcessId)
    $currentThreadId = $Win32::GetCurrentThreadId()
    $foregroundBefore = $Win32::GetForegroundWindow()
    $foregroundProcessId = [uint32]0
    $foregroundThreadId = $Win32::GetWindowThreadProcessId($foregroundBefore, [ref]$foregroundProcessId)
    $foregroundAttached = $false
    $targetAttached = $false
    $windowFlags = 0x0001 -bor 0x0002 -bor 0x0040 # SWP_NOSIZE | SWP_NOMOVE | SWP_SHOWWINDOW
    try {
        # First promote the VBE from the current foreground thread. This is the
        # Windows-supported path around foreground-lock restrictions.
        if ($foregroundThreadId -and $foregroundThreadId -ne $currentThreadId) {
            $foregroundAttached = $Win32::AttachThreadInput($currentThreadId, $foregroundThreadId, $true)
        }
        [void]$Win32::ShowWindow($hwnd, 9) # SW_RESTORE
        [void]$Win32::SetWindowPos($hwnd, [IntPtr](-1), 0, 0, 0, 0, $windowFlags) # HWND_TOPMOST
        [void]$Win32::BringWindowToTop($hwnd)
        [void]$Win32::SetForegroundWindow($hwnd)

        $foreground = $Win32::GetForegroundWindow()
        if ($foreground -ne $hwnd) {
            throw "VBE was not brought to the foreground. Expected 0x$($hwnd.ToString('X8')), got 0x$($foreground.ToString('X8'))."
        }
    } finally {
        if ($foregroundAttached) {
            [void]$Win32::AttachThreadInput($currentThreadId, $foregroundThreadId, $false)
        }
    }

    try {
        # Then give the PowerShell thread access to the VBE input queue so
        # SetFocus can establish keyboard focus inside the already-foreground frame.
        if ($targetThreadId -and $targetThreadId -ne $currentThreadId) {
            $targetAttached = $Win32::AttachThreadInput($currentThreadId, $targetThreadId, $true)
        }
        [void]$Win32::SetActiveWindow($hwnd)
        [void]$Win32::SetFocus($hwnd)
        $focused = $Win32::GetFocus()
        $focusedRoot = $Win32::GetAncestor($focused, 2) # GA_ROOT
        if ($focused -eq [IntPtr]::Zero -or $focusedRoot -ne $hwnd) {
            throw "SetFocus did not focus the VBE window. Expected root 0x$($hwnd.ToString('X8')), got focus 0x$($focused.ToString('X8')) with root 0x$($focusedRoot.ToString('X8'))."
        }
        $foreground = $Win32::GetForegroundWindow()
        if ($foreground -ne $hwnd) {
            throw "VBE was not the foreground window. Expected 0x$($hwnd.ToString('X8')), got 0x$($foreground.ToString('X8'))."
        }
        if (-not [W32Input]::SendAltF4()) {
            throw "SendInput failed while sending Alt+F4 to VBE HWND 0x$($hwnd.ToString('X8'))."
        }
        [void]$Win32::SetWindowPos($hwnd, [IntPtr](-2), 0, 0, 0, 0, $windowFlags) # HWND_NOTOPMOST
    } finally {
        if ($targetAttached) {
            [void]$Win32::AttachThreadInput($currentThreadId, $targetThreadId, $false)
        }
    }
    Start-Sleep -Milliseconds 300

    $excelAfterAltQ = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    try {
        if (-not [bool]$excelAfterAltQ.VBE.MainWindow.Visible) {
            Write-Host "  VBE returned to Excel through verified Alt+F4 (MainWindow.Visible=False)." -ForegroundColor Yellow
            return
        }
    } finally {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelAfterAltQ) | Out-Null
    }

    for ($i = 0; $i -lt 20; $i++) {
        Start-Sleep -Milliseconds 250
        if ((Find-VBEWindow) -eq [IntPtr]::Zero) {
            Write-Host "  VBE closed through verified Alt+F4." -ForegroundColor Yellow
            return
        }
    }
    throw "VBE window did not close after verified Alt+F4."
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

    Step "2. Prepare the vbapm project"
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
codename = "GhostDemoProject"

[source.files]
Module1 = "src/Module1.bas"
"@ | Set-Content (Join-Path $projectDir "vbaproject.toml")
    Write-Host "Project initialized." -ForegroundColor Green

    Step "3. Build the project"
    Run-Cmd "vba build" $projectDir
    Write-Host "Project built." -ForegroundColor Green

    Step "4. Open new visible Excel instance"
    & (Join-Path $scripts "Create-NewExcel.ps1")
    # Wait until Excel registers in the ROT
    $excel = $null
    for ($i = 0; $i -lt 10 -and -not $excel; $i++) {
        try { $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") }
        catch { Start-Sleep -Seconds 1 }
    }
    if (-not $excel) { throw "Excel never became COM-reachable after open." }
    $excelHwnd = [IntPtr]$excel.Hwnd
    $script:ExcelProcessId = [uint32]0
    $Win32::GetWindowThreadProcessId($excelHwnd, [ref]$script:ExcelProcessId) | Out-Null
    if (-not $script:ExcelProcessId) { throw "Could not determine the target Excel process ID." }
    Start-Sleep -Seconds 3
    Write-Host "Excel ready." -ForegroundColor Green

    Step "5. Open VBE window"
    $vbeOpen = Open-VBE

    Step "6. Open the built file"
    # Open via COM (Workbooks.Open) instead of `vba open`.
    # `vba open` shells out to `Start -Wait` which hangs in headless CI.
    $builtFile = Join-Path $projectDir "build/ghost-demo.xlsm"
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

    Write-Host "VBE projects after workbook open:" -ForegroundColor Cyan
    & (Join-Path $scripts "detect-ghost-vbe.ps1")

    Step "7. Close the workbook"
    if (($HideVBEBeforeClose -or $CloseVBEBeforeClose) -and $vbeOpen) {
        $excelForClose = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        try {
            if ($CloseVBEBeforeClose) {
                $sender = Join-Path $scripts "Send-AltQ-To-VBE.ps1"
                if (!(Test-Path $sender)) { throw "Alt+F4 sender not found at $sender" }
                & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $sender -Key F4 -ProcessId $script:ExcelProcessId
                if ($LASTEXITCODE -ne 0) { throw "Alt+F4 sender failed (exit $LASTEXITCODE)." }
                if ([bool]$excelForClose.VBE.MainWindow.Visible) {
                    throw "Alt+F4 sender completed but VBE.MainWindow.Visible is still True."
                }
                Write-Host "VBE returned to Excel through verified Alt+F4 (MainWindow.Visible=False)." -ForegroundColor Yellow
            } else {
                $excelForClose.VBE.MainWindow.Visible = $false
                Write-Host "VBE hidden before workbook close." -ForegroundColor Yellow
            }
        } finally {
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelForClose) | Out-Null
        }
    }
    $wbName = "ghost-demo.xlsm"
    if ($UseCliClose) {
        $installedVba = Join-Path $env:APPDATA "vbapm/bin/vba.cmd"
        if (!(Test-Path $installedVba)) { throw "Installed vba command not found at $installedVba" }
        Run-Cmd "& '$installedVba' close --force" $projectDir
    } else {
        & (Join-Path $scripts "close-workbook.ps1") -Name $wbName -Force
    }
    if ($UseCliClose) {
        Write-Host "Workbook closed through vbapm VBA (Build.CloseFile)." -ForegroundColor Green
    } else {
        Write-Host "Workbook closed through the standalone COM helper." -ForegroundColor Green
    }

    Step "8. Check for ghost VBE projects"
    $detectionJson = & (Join-Path $scripts "detect-ghost-vbe.ps1") -Quiet | Out-String
    $detectionSummary = $detectionJson.Trim() | ConvertFrom-Json
    $ghostCount = [int]$detectionSummary.Ghosts
    Write-Output "GHOST_COUNT_AFTER_WORKBOOK_CLOSE=$ghostCount"

    if ($CloseAndReopenVBEAfterClose) {
        Step "9. Close and reopen VBE after workbook close"
        Close-VBEWindow
        if (-not (Open-VBE)) { throw "VBE did not reopen after WM_CLOSE." }

        Step "10. Check for ghosts after VBE reopen"
        Write-Host ""
        & (Join-Path $scripts "detect-ghost-vbe.ps1")
        $reopenDetectionJson = & (Join-Path $scripts "detect-ghost-vbe.ps1") -Quiet | Out-String
        $reopenDetectionSummary = $reopenDetectionJson.Trim() | ConvertFrom-Json
        $ghostCount = [int]$reopenDetectionSummary.Ghosts
        Write-Output "GHOST_COUNT_AFTER_VBE_REOPEN=$ghostCount"
    }

    if ($RefreshVBEAfterClose) {
        Step "9. Refresh VBE after workbook close"
        $excelAfterClose = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        try {
            $excelAfterClose.VBE.MainWindow.Visible = $false
            Start-Sleep -Milliseconds 500
            $excelAfterClose.VBE.MainWindow.Visible = $true
            Start-Sleep -Seconds 2
            Write-Host "  VBE visibility toggled after workbook close." -ForegroundColor Yellow
        } finally {
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelAfterClose) | Out-Null
        }

        Step "10. Check for ghosts after VBE refresh"
        $refreshDetectionJson = & (Join-Path $scripts "detect-ghost-vbe.ps1") -Quiet | Out-String
        $refreshDetectionSummary = $refreshDetectionJson.Trim() | ConvertFrom-Json
        $ghostCount = [int]$refreshDetectionSummary.Ghosts
        Write-Output "GHOST_COUNT_AFTER_VBE_REFRESH=$ghostCount"
    }

    Write-Output "GHOST_COUNT=$ghostCount"

    if ($HoldAfterDetection -and $ghostCount -gt 0) {
        Step "Manual inspection hold"
        $inspectionExcel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        $inspectionVbe = $inspectionExcel.VBE
        $inspectionProjects = @($inspectionVbe.VBProjects)
        Write-Host "Holding Excel/VBE COM references for manual inspection. Press Ctrl+C in this terminal when finished." -ForegroundColor Yellow
        foreach ($inspectionProject in $inspectionProjects) {
            try { Write-Host ("  VBE project: {0} | FileName={1}" -f $inspectionProject.Name, $inspectionProject.FileName) } catch { }
        }
        while ($true) { Start-Sleep -Seconds 5 }
    }

    Write-Host ""; Write-Host "===== Repro complete =====" -ForegroundColor Cyan
} finally { Pop-Location }
