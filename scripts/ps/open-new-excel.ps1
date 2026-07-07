<#
.SYNOPSIS
    Opens a new, visible Excel instance as a separate process.

.DESCRIPTION
    Uses Start-Process with the /x flag to force a genuinely new
    EXCEL.EXE process — never attaches to an existing instance.
    Useful for getting a clean instance that vbapm can target.

.PARAMETER File
    Path to a workbook to open in the new instance.

.PARAMETER Blank
    Open with a new blank workbook (default if no File is given).

.EXAMPLE
    .\scripts\ps\open-new-excel.ps1

.EXAMPLE
    .\scripts\ps\open-new-excel.ps1 -File "C:\my-project\build\my-project.xlsm"
#>

param(
    [string]$File,
    [switch]$Blank
)

# No COM needed for Start-Process, so no PS Core redirect necessary.
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
} else {
    Write-Host "Opening with a new blank workbook" -ForegroundColor Yellow
}

$proc = Start-Process excel.exe -ArgumentList $args -PassThru
Write-Host "New instance started: PID $($proc.Id)" -ForegroundColor Green

# Give Excel a moment to register in the ROT
Start-Sleep -Seconds 2
Write-Host ""
Write-Host "Run .\scripts\ps\count-excel-instances.ps1 to verify both instances."
