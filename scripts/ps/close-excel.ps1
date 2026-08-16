<#
.SYNOPSIS
    Closes the COM-reachable (ROT) Excel instance, but refuses if any
    workbook has unsaved changes.

.PARAMETER Force
    Quit even if there are unsaved workbooks (changes will be lost).

.PARAMETER SaveAll
    Save all unsaved workbooks automatically, then quit.

.PARAMETER List
    Just show open workbooks and their saved status without quitting.

.EXAMPLE
    .\scripts\ps\close-excel.ps1

.EXAMPLE
    .\scripts\ps\close-excel.ps1 -List

.EXAMPLE
    .\scripts\ps\close-excel.ps1 -Force

.EXAMPLE
    .\scripts\ps\close-excel.ps1 -SaveAll
#>

param(
    [switch]$Force,
    [switch]$SaveAll,
    [switch]$List
)

# ---------------------------------------------------------------
# PowerShell Core -> re-invoke with Windows PowerShell for COM
# ---------------------------------------------------------------
if ($PSVersionTable.PSEdition -eq 'Core') {
    $winPS = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($winPS) {
        $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
        if ($Force)   { $psArgs += '-Force' }
        if ($SaveAll) { $psArgs += '-SaveAll' }
        if ($List)    { $psArgs += '-List' }
        & powershell.exe $psArgs
        exit $LASTEXITCODE
    }
}

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------
# Attach
# ---------------------------------------------------------------
try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    Write-Host "No COM-reachable Excel instance running." -ForegroundColor Green
    exit 0
}

# ---------------------------------------------------------------
# Check all workbooks
# ---------------------------------------------------------------
$unsaved  = @()
$allWbs   = @()

foreach ($wb in $excel.Workbooks) {
    $entry = [PSCustomObject]@{
        Name     = $wb.Name
        Path     = $wb.FullName
        Saved    = $wb.Saved
    }
    $allWbs += $entry
    if (-not $wb.Saved) {
        $unsaved += $entry
    }
}

# Also check add-in workbooks.  They may not appear in the
# Workbooks enumeration but can be accessed by name directly.
foreach ($ai in $excel.AddIns) {
    if (-not $ai.Installed) { continue }
    # Already counted if it appeared in the Workbooks enumeration
    $alreadySeen = $false
    foreach ($wb in $allWbs) { if ($wb.Name -eq $ai.Name) { $alreadySeen = $true; break } }
    if ($alreadySeen) { continue }

    try {
        $wb = $excel.Workbooks($ai.Name)
        if ($wb) {
            $entry = [PSCustomObject]@{ Name = $wb.Name; Path = $wb.FullName; Saved = $wb.Saved }
            $allWbs += $entry
            if (-not $wb.Saved) { $unsaved += $entry }
        }
    } catch { }
}
Write-Host "===== Workbooks =====" -ForegroundColor Cyan
Write-Host "Total: $($allWbs.Count)  |  Unsaved: $($unsaved.Count)"
Write-Host ""

foreach ($wb in $allWbs) {
    $tag = ""
    try { if ($excel.Workbooks($wb.Name).IsAddin) { $tag = " [add-in]" } } catch { }
    if ($wb.Saved) {
        Write-Host "  [saved]   $($wb.Name)$tag" -ForegroundColor Green
    } else {
        Write-Host "  [MODIFIED] $($wb.Name)$tag" -ForegroundColor Red -NoNewline
        Write-Host "  $($wb.Path)"
    }
}

if ($allWbs.Count -eq 0) {
    Write-Host ""
    Write-Host "No workbooks open. Quitting..." -ForegroundColor Yellow
    # Suppress any "save add-in?" dialogs
    $excel.DisplayAlerts = $false
    $excel.Quit()
    Write-Host "Excel closed." -ForegroundColor Green
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 0
}

# ---------------------------------------------------------------
# List-only mode
# ---------------------------------------------------------------
if ($List) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 0
}

# ---------------------------------------------------------------
# Decide
# ---------------------------------------------------------------
if ($unsaved.Count -gt 0 -and -not $Force -and -not $SaveAll) {
    Write-Host ""
    Write-Host "Refusing to quit: $($unsaved.Count) workbook(s) have unsaved changes." -ForegroundColor Red
    Write-Host "Use -SaveAll to save and quit, -Force to discard, or save them first."
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 1
}

# ---------------------------------------------------------------
# Save all if requested
# ---------------------------------------------------------------
if ($SaveAll -and $unsaved.Count -gt 0) {
    Write-Host ""
    Write-Host "Saving $($unsaved.Count) unsaved workbook(s)..." -ForegroundColor Yellow
    foreach ($entry in $unsaved) {
        try {
            $wb = $excel.Workbooks($entry.Name)
            $wb.Save()
            Write-Host "  Saved: $($entry.Name)" -ForegroundColor Green
        } catch {
            Write-Host "  FAILED to save: $($entry.Name) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# ---------------------------------------------------------------
# Quit
# ---------------------------------------------------------------
if ($Force -and $unsaved.Count -gt 0 -and -not $SaveAll) {
    Write-Host ""
    Write-Host "Force-quitting: $($unsaved.Count) unsaved workbook(s) will lose changes." -ForegroundColor Yellow
}

$excel.DisplayAlerts = $false
$excel.Quit()
Write-Host "Excel closed." -ForegroundColor Green
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
