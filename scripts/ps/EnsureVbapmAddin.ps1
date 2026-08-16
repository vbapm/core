# EnsureVbapmAddin.ps1
#
# Ensure the vbapm add-in (vbapm.xlam) is open in a (visible) Excel instance
# before running the e2e suite, and verify it points at the freshly-built repo
# add-in (not a stale %APPDATA% copy).
#
# Behavior:
#   1. Attach to the running visible Excel instance via GetActiveObject, or
#      launch a new visible one if none is running.
#   2. If vbapm.xlam is not open, open it from <repo>/addins/build/vbapm.xlam.
#   3. If it IS already open, verify its FullName resolves to the repo build
#      path (identity check) — reporting a warning if it points elsewhere (e.g.
#      a stale installed copy).
#   4. Print the sha256 of the on-disk built add-in for reference.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File `
#     .\scripts\ps\EnsureVbapmAddin.ps1

[CmdletBinding()]
param(
    # Absolute path to the built vbapm.xlam (addins/build/vbapm.xlam).
    [string]$AddinPath = ''
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

if (-not $AddinPath) {
    $AddinPath = Join-Path $PSScriptRoot '..\..\addins\build\vbapm.xlam'
}
$AddinPath = [System.IO.Path]::GetFullPath($AddinPath)
$AddinFile = [System.IO.Path]::GetFileName($AddinPath)

# 1) Attach to (or create) a visible Excel instance.
$app = $null
$appWasOpen = $false
try {
    $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $appWasOpen = $true
} catch {
    $app = $null
}
if ($null -eq $app) {
    $app = New-Object -ComObject "Excel.Application"
    $app.Visible = $true
}

# 2/3) Ensure vbapm.xlam is open and matches the repo build path.
$openAddin = $null
foreach ($wb in $app.Workbooks) {
    try {
        if ((Get-Item $wb.FullName).FullName -eq $AddinPath) {
            $openAddin = $wb
            break
        }
    } catch {
        # workbook without a resolvable path; skip
    }
}

if ($null -eq $openAddin) {
    # Check whether a DIFFERENT vbapm.xlam is already open (stale copy).
    $otherOpen = $null
    foreach ($wb in $app.Workbooks) {
        try {
            if ([System.IO.Path]::GetFileName($wb.FullName) -eq $AddinFile) {
                $otherOpen = $wb.FullName
                break
            }
        } catch {}
    }
    if ($otherOpen) {
        Write-Warning "A vbapm.xlam is already open from: $otherOpen"
        Write-Warning "Expected the repo build at: $AddinPath"
        Write-Warning "Close the stale copy to ensure e2e uses the freshly-built add-in."
    } else {
        $null = $app.Workbooks.Open($AddinPath)
        Write-Output "Opened vbapm.xlam from $AddinPath"
    }
} else {
    Write-Output "vbapm.xlam already open from the repo build: $AddinPath"
}

# 4) Print sha256 of the on-disk build for reference (best-effort — the file is
#    locked once Excel opens it, so hash only succeeds before first open).
try {
    $hash = (Get-FileHash -Path $AddinPath -Algorithm SHA256 -ErrorAction SilentlyContinue).Hash
    if ($hash) {
        Write-Output "Built vbapm.xlam sha256: $($hash.ToLowerInvariant())"
    }
} catch {
    # ignore — the add-in may already be open (file locked).
}

# Register/refresh the instance in the coordination registry.
if (Get-Command Register-ExcelInstance -ErrorAction SilentlyContinue) {
    try {
        Register-ExcelInstance -ExcelApp $app -Owner "e2e-setup#$PID" -Visible $true -Reason 'e2e-setup' | Out-Null
    } catch {
        # best-effort
    }
}

if (-not $appWasOpen) {
    Write-Output "Launched a new visible Excel instance. It will stay open for the e2e run."
} else {
    Write-Output "Attached to an already-running visible Excel instance."
}
