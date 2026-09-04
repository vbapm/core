<#
.SYNOPSIS
    Unloads a specific Excel add-in (e.g. vbapm.xlam) from the COM-reachable
    (ROT) Excel instance, releasing its file lock without quitting Excel or
    touching other open workbooks.

.PARAMETER Name
    Add-in name to unload (e.g. "vbapm").  Matches against the add-in's .Name
    property (case-insensitive).  The ".xlam" extension is optional.

.PARAMETER Path
    Full path to the add-in file to unload.  Matches against the workbook's
    .FullName for add-ins that are open as workbooks.

.PARAMETER List
    Just list installed/loaded add-ins without unloading anything.

.EXAMPLE
    .\scripts\ps\close-addin.ps1 -List

.EXAMPLE
    .\scripts\ps\close-addin.ps1 -Name "vbapm"

.EXAMPLE
    .\scripts\ps\close-addin.ps1 -Path "C:\Users\me\AppData\Roaming\vbapm\addins\build\vbapm.xlam"
#>

param(
    [string]$Name,
    [string]$Path,
    [switch]$List
)

# ---------------------------------------------------------------
# PowerShell Core -> re-invoke with Windows PowerShell for COM
# ---------------------------------------------------------------
if ($PSVersionTable.PSEdition -eq 'Core') {
    $winPS = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($winPS) {
        $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
        if ($Name) { $psArgs += '-Name'; $psArgs += $Name }
        if ($Path) { $psArgs += '-Path'; $psArgs += $Path }
        if ($List) { $psArgs += '-List' }
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
    Write-Host "Add-ins in ROT instance:" -ForegroundColor Cyan
    foreach ($ai in $excel.AddIns) {
        Write-Host "  $($ai.Name)  installed=$($ai.Installed) open=$($ai.IsOpen)"
    }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 0
}

if (-not $Name -and -not $Path) {
    Write-Error "Specify -Name or -Path of the add-in to unload (use -List to see installed add-ins)."
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 1
}

# Normalize the requested name (strip optional .xlam/.xlsm extension)
$targetName = $null
if ($Name) {
    $targetName = $Name
    if ($targetName -match '(?i)\.(xlam|xlsm)$') {
        $targetName = [System.IO.Path]::GetFileNameWithoutExtension($targetName)
    }
}

# ---------------------------------------------------------------
# Unload via the AddIns collection
# ---------------------------------------------------------------
$unloaded = $false
foreach ($ai in $excel.AddIns) {
    $matches = $false
    if ($targetName -and $ai.Name -eq $targetName) {
        $matches = $true
    }
    if ($Path) {
        try {
            $wb = $excel.Workbooks($ai.Name)
            if ($wb -and [System.IO.Path]::GetFullPath($wb.FullName) -eq [System.IO.Path]::GetFullPath($Path)) {
                $matches = $true
            }
        } catch {
            # add-in not open as a workbook; skip
        }
    }

    if ($matches) {
        Write-Host "Unloading add-in: $($ai.Name) (installed=$($ai.Installed))" -ForegroundColor Yellow
        if ($ai.Installed) {
            $ai.Installed = $false
        }
        $unloaded = $true
    }
}

# ---------------------------------------------------------------
# Also try a direct workbook close by name.  A previously opened
# .xlam can linger with an open file handle even when it no longer
# appears in the Workbooks/AddIns enumerations.
# ---------------------------------------------------------------
if (-not $unloaded) {
    $nameToTry = $targetName
    if (-not $nameToTry -and $Path) {
        $nameToTry = [System.IO.Path]::GetFileName($Path)
    }
    if ($nameToTry) {
        foreach ($candidate in @($nameToTry, "$nameToTry.xlam", "$nameToTry.xlsm")) {
            try {
                $wb = $excel.Workbooks($candidate)
                if ($wb) {
                    Write-Host "Closing workbook: $($wb.Name) [$($wb.FullName)]" -ForegroundColor Yellow
                    $wb.Close($false)
                    $unloaded = $true
                }
            } catch {
                # not found under this name; try next
            }
        }
    }
}

if (-not $unloaded) {
    Write-Warning "Add-in not found in ROT instance.  Use -List to see installed add-ins."
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 1
}

Write-Host "Add-in unloaded." -ForegroundColor Green
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
