<#
.SYNOPSIS
    Closes a specific workbook in the COM-reachable (ROT) Excel instance.

.PARAMETER Name
    Workbook name to close (e.g. "Book1" or "Book1.xlsm").  Matches
    against the workbook's .Name property (case-insensitive).

.PARAMETER Path
    Full path to the workbook to close.  Matches against .FullName.

.PARAMETER Save
    Save changes before closing.  Default is to discard.

.PARAMETER List
    Just list open workbooks without closing anything.

.PARAMETER Force
    Skip the confirmation prompt.

.EXAMPLE
    .\scripts\ps\close-workbook.ps1 -List

.EXAMPLE
    .\scripts\ps\close-workbook.ps1 -Name "standard.xlsm" -Force

.EXAMPLE
    .\scripts\ps\close-workbook.ps1 -Path "D:\my-project\build\my-project.xlsm" -Save
#>

param(
    [string]$Name,
    [string]$Path,
    [switch]$Save,
    [switch]$List,
    [switch]$Force
)

# ---------------------------------------------------------------
# PowerShell Core -> re-invoke with Windows PowerShell for COM
# ---------------------------------------------------------------
if ($PSVersionTable.PSEdition -eq 'Core') {
    $winPS = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($winPS) {
        $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
        if ($Name)  { $psArgs += '-Name'; $psArgs += $Name }
        if ($Path)  { $psArgs += '-Path'; $psArgs += $Path }
        if ($Save)  { $psArgs += '-Save' }
        if ($List)  { $psArgs += '-List' }
        if ($Force) { $psArgs += '-Force' }
        & powershell.exe $psArgs
        exit $LASTEXITCODE
    }
}

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------
# Attach to ROT instance
# ---------------------------------------------------------------
try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    Write-Error "No COM-reachable Excel instance.  Start Excel or check for zombie processes."
    exit 1
}

# ---------------------------------------------------------------
# List mode
# ---------------------------------------------------------------
if ($List) {
    Write-Host "Workbooks open in ROT instance:" -ForegroundColor Cyan
    foreach ($wb in $excel.Workbooks) {
        $saved = if ($wb.Saved) { "saved" } else { "MODIFIED" }
        Write-Host "  $($wb.Name)  [$saved]  $($wb.FullName)"
    }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 0
}

# ---------------------------------------------------------------
# Find the workbook
# ---------------------------------------------------------------
if (-not $Name -and -not $Path) {
    Write-Error "Specify -Name or -Path of the workbook to close (use -List to see open workbooks)."
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 1
}

$target = $null
$workbooks = $null
try {
    $workbooks = $excel.Workbooks
    $normalizedTarget = if ($Path) { [System.IO.Path]::GetFullPath($Path) } else { $null }
    for ($index = 1; $index -le $workbooks.Count; $index++) {
        $wb = $workbooks.Item($index)
        try {
            if ($Name -and $wb.Name -eq $Name) {
                $target = $wb
                $wb = $null
                break
            }
            if ($normalizedTarget) {
                try {
                    $normalizedWb = [System.IO.Path]::GetFullPath($wb.FullName)
                    if ($normalizedWb -eq $normalizedTarget) {
                        $target = $wb
                        $wb = $null
                        break
                    }
                } catch { }
            }
        } finally {
            if ($null -ne $wb) {
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null
            }
        }
    }
} finally {
    if ($null -ne $workbooks) {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbooks) | Out-Null
    }
}

if (-not $target) {
    Write-Error "Workbook not found in ROT instance.  Use -List to see open workbooks."
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 1
}

# ---------------------------------------------------------------
# Confirm & close
# ---------------------------------------------------------------
$saved   = if ($target.Saved) { "saved" } else { "MODIFIED" }
$wbName  = $target.Name
$wbPath  = $target.FullName

Write-Host "Workbook: $wbName  [$saved]" -ForegroundColor Yellow
Write-Host "Path:      $wbPath"
Write-Host "Action:    $(if ($Save) { 'Save + close' } else { 'Close (discard changes)' })"
Write-Host ""

if (-not $Force) {
    $confirm = Read-Host "Proceed? [y/N]"
    if ($confirm -notmatch '^[yY]') {
        Write-Host "Aborted." -ForegroundColor Yellow
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        exit 0
    }
}

$saveChanges = [bool]$Save
try {
    $target.Close($saveChanges)
    Write-Host "Closed: $wbName" -ForegroundColor Green
} finally {
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($target) | Out-Null
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
}
