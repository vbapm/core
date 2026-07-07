<#
.SYNOPSIS
    Opens a new, visible Excel instance as a separate process.

.DESCRIPTION
    Uses Start-Process with the /x flag to force a genuinely new
    EXCEL.EXE process — never attaches to an existing instance.
    By default, creates a blank workbook so you're not staring at
    the start screen.

.PARAMETER File
    Path to a workbook to open in the new instance.

.PARAMETER NoBlank
    Don't create a blank workbook (leave Excel at the start screen).

.EXAMPLE
    .\scripts\ps\open-new-excel.ps1

.EXAMPLE
    .\scripts\ps\open-new-excel.ps1 -File "C:\my-project\build\my-project.xlsm"

.EXAMPLE
    .\scripts\ps\open-new-excel.ps1 -NoBlank
#>

param(
    [string]$File,
    [switch]$NoBlank
)

# ---------------------------------------------------------------
# PowerShell Core -> re-invoke with Windows PowerShell for COM
# ---------------------------------------------------------------
if ($PSVersionTable.PSEdition -eq 'Core') {
    $winPS = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($winPS) {
        $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
        if ($File)    { $psArgs += '-File'; $psArgs += $File }
        if ($NoBlank) { $psArgs += '-NoBlank' }
        & powershell.exe $psArgs
        exit $LASTEXITCODE
    }
}

$ErrorActionPreference = 'Stop'

Write-Host "Launching new Excel process..." -ForegroundColor Cyan

$args = @('/x')

if ($File) {
    if (-not (Test-Path $File)) {
        Write-Error "File not found: $File"
        exit 1
    }
    $resolved = (Resolve-Path $File).Path
    $args += "`"$resolved`""
    Write-Host "Opening: $resolved" -ForegroundColor Yellow
}

$proc = Start-Process excel.exe -ArgumentList $args -PassThru
Write-Host "New instance started: PID $($proc.Id)" -ForegroundColor Green

# Give Excel time to initialize and register in ROT
Start-Sleep -Seconds 3

# Create a blank workbook so we're not at the start screen
if (-not $File -and -not $NoBlank) {
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        $excel.Workbooks.Add() | Out-Null
        Write-Host "Blank workbook created." -ForegroundColor Green
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    } catch {
        Write-Host "(could not create blank workbook via COM - Excel may not be ready yet)" -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "Run .\scripts\ps\count-excel-instances.ps1 to verify."
