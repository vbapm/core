<#
.SYNOPSIS
    Creates a new Excel instance as a separate process.

.DESCRIPTION
    Uses COM activation to create a genuinely new Excel.Application object.
    Workbook operations use that returned object directly, so they cannot be
    redirected to another Excel instance. Use -Invisible to hide the new
    instance. By default, creates a blank workbook so you're not staring at
    the start screen.

.PARAMETER File
    Path to a workbook to open in the new instance.

.PARAMETER NoBlank
    Don't create a blank workbook (leave Excel at the start screen).

.PARAMETER Invisible
    Start Excel without showing its window.

.EXAMPLE
    .\scripts\ps\Create-NewExcel.ps1

.EXAMPLE
    .\scripts\ps\Create-NewExcel.ps1 -File "C:\my-project\build\my-project.xlsm"

.EXAMPLE
    .\scripts\ps\Create-NewExcel.ps1 -NoBlank

.EXAMPLE
    .\scripts\ps\Create-NewExcel.ps1 -Invisible
#>

param(
    [string]$File,
    [switch]$NoBlank,
    [switch]$Invisible
)

if ($File -eq '--invisible') {
    $File = $null
    $Invisible = $true
}

# ---------------------------------------------------------------
# PowerShell Core -> re-invoke with Windows PowerShell for COM
# ---------------------------------------------------------------
if ($PSVersionTable.PSEdition -eq 'Core') {
    $winPS = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($winPS) {
        $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
        if ($File)    { $psArgs += '-File'; $psArgs += $File }
        if ($NoBlank) { $psArgs += '-NoBlank' }
        if ($Invisible) { $psArgs += '-Invisible' }
        & powershell.exe $psArgs
        exit $LASTEXITCODE
    }
}

$ErrorActionPreference = 'Stop'

if (-not ('CreateNewExcelWindow' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class CreateNewExcelWindow {
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr windowHandle, int command);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(
        IntPtr windowHandle,
        out uint processId);
}
"@
}

function Get-CreatedExcelProcessId {
    param([object]$ExcelApp)

    [uint32]$processId = 0
    [CreateNewExcelWindow]::GetWindowThreadProcessId([IntPtr]$ExcelApp.Hwnd, [ref]$processId) | Out-Null
    if ($processId -gt 0) {
        return [int]$processId
    }

    return $null
}

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

function Set-ExcelVisibility {
    param(
        [object]$ExcelApp,
        [bool]$Visible
    )

    $ExcelApp.Visible = $Visible
    if (-not $Visible -and $ExcelApp.Hwnd) {
        [CreateNewExcelWindow]::ShowWindow([IntPtr]$ExcelApp.Hwnd, 0) | Out-Null
    }
}

Write-Host "Launching new Excel process..." -ForegroundColor Cyan

$resolved = $null

if ($File) {
    if (-not (Test-Path $File)) {
        Write-Error "File not found: $File"
        exit 1
    }
    $resolved = (Resolve-Path $File).Path
    Write-Host "Opening: $resolved" -ForegroundColor Yellow
}

$excel = $null
try {
    $excelType = [System.Type]::GetTypeFromProgID('Excel.Application')
    if ($null -eq $excelType) {
        throw 'Excel.Application is not registered on this machine.'
    }

    $excel = [System.Activator]::CreateInstance($excelType)
    Set-ExcelVisibility -ExcelApp $excel -Visible (-not $Invisible)

    $processId = Get-CreatedExcelProcessId -ExcelApp $excel
    if ($processId) {
        Write-Host "New instance started: PID $processId" -ForegroundColor Green
    } else {
        Write-Host 'New instance started: PID unknown' -ForegroundColor Green
    }

    if ($resolved) {
        $workbook = $excel.Workbooks.Open($resolved)
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) | Out-Null
        Write-Host "Workbook opened: $resolved" -ForegroundColor Green
    } elseif (-not $NoBlank) {
        $workbook = $excel.Workbooks.Add()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) | Out-Null
        Write-Host "Blank workbook created." -ForegroundColor Green
    }

    if ($Invisible) {
        Set-ExcelVisibility -ExcelApp $excel -Visible $false
    }

    $registeredPid = Register-ExcelInstance -ExcelApp $excel -Owner "terminal-$PID" -Reason 'Create-NewExcel'
    if ($registeredPid) {
        Write-Host "Instance registered: PID $registeredPid" -ForegroundColor DarkCyan
    }
} finally {
    if ($null -ne $excel) {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
}

Write-Host ""
Write-Host "Run .\scripts\ps\count-excel-instances.ps1 to verify."
