<#
.SYNOPSIS
    Detects ghost VBA projects in the VBE — projects visible in the VBE
    whose parent workbook has already been closed.

.DESCRIPTION
    When an Excel workbook is closed (especially after COM automation has
    interacted with its VBProject), the VBE sometimes fails to clean up
    its internal reference.  The project remains listed in the VBE Project
    Explorer even though the workbook is gone.  This script identifies
    those "ghost" projects by cross-referencing VBE.VBProjects against
    Application.Workbooks.

.PARAMETER Name
    If provided, only check for a ghost project with this name.

.PARAMETER Quiet
    Suppress per-project output; only print JSON summary at the end.

.PARAMETER Raw
    Output a raw JSON array so the result can be consumed by another tool.

.EXAMPLE
    .\scripts\ps\detect-ghost-vbe.ps1

.EXAMPLE
    .\scripts\ps\detect-ghost-vbe.ps1 -Name "VBAProject"

.EXAMPLE
    .\scripts\ps\detect-ghost-vbe.ps1 -Raw
#>

param(
    [string]$Name,
    [switch]$Quiet,
    [switch]$Raw
)

# ---------------------------------------------------------------
# PowerShell Core detection — re-invoke with Windows PowerShell
# ---------------------------------------------------------------
# pwsh (PowerShell 7+) runs on .NET (Core) which may lack the
# Marshal.GetActiveObject API needed for COM attachment.  Windows
# PowerShell 5.1 (powershell.exe) uses .NET Framework where the
# full COM interop surface is available.  Re-invoke transparently.
if ($PSVersionTable.PSEdition -eq 'Core') {
    $winPS = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($winPS) {
        $psArgs = @(
            '-NoProfile',
            '-ExecutionPolicy', 'Bypass',
            '-File', $PSCommandPath
        )
        if ($Name)  { $psArgs += '-Name'; $psArgs += $Name }
        if ($Quiet) { $psArgs += '-Quiet' }
        if ($Raw)   { $psArgs += '-Raw' }
        & powershell.exe $psArgs
        exit $LASTEXITCODE
    }
}

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------
# Step 1 — attach to the running Excel instance
# ---------------------------------------------------------------
# Quick pre-check: is an Excel process even running?
$excelProcs = @(Get-Process excel -ErrorAction SilentlyContinue)
if ($excelProcs.Count -eq 0) {
    $err = @{
        error = "No Excel process found. Start Excel and try again."
    }
    if ($Raw) { ConvertTo-Json $err; exit 1 } else { Write-Error $err.error; exit 1 }
}

$Excel = $null
$lastError = ""

# Strategy A: Marshal.GetActiveObject (available in .NET Framework)
try {
    $Excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    $lastError = "GetActiveObject: $($_.Exception.Message)"
}

# Strategy B: versioned ProgIDs
if (-not $Excel) {
    foreach ($suffix in @("16", "15")) {
        try {
            $Excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application.$suffix")
            if ($Excel) { break }
        } catch { }
    }
}

# Strategy C: VB GetObject
if (-not $Excel) {
    try {
        Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
        $Excel = [Microsoft.VisualBasic.Interaction]::GetObject($null, "Excel.Application")
    } catch {
        $lastError += " | VB GetObject: $($_.Exception.Message)"
    }
}

# Strategy D: New-Object (creates fresh instance — last resort)
if (-not $Excel) {
    try {
        $Excel = New-Object -ComObject "Excel.Application"
        if ($Excel.Workbooks.Count -eq 0 -and $excelProcs.Count -gt 0) {
            $Excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
            $Excel = $null
        }
    } catch {
        $lastError += " | New-Object: $($_.Exception.Message)"
    }
}

if (-not $Excel) {
    $err = @{
        error = "Cannot attach to a running Excel instance."
        detail = $lastError
        hint   = "Ensure Excel is running at the same elevation level as this terminal."
    }
    if ($Raw) { ConvertTo-Json $err -Depth 2; exit 1 } else { Write-Error $err.error; Write-Host $err.hint -ForegroundColor Yellow; exit 1 }
}

# ---------------------------------------------------------------
# Step 2 — build a set of paths for every open workbook / add-in
# ---------------------------------------------------------------
$OpenPaths = @{}     # key = normalized full path, value = display name

# Workbooks
foreach ($wb in $Excel.Workbooks) {
    try {
        $path = [System.IO.Path]::GetFullPath($wb.FullName)
        $OpenPaths[$path.ToLowerInvariant()] = $wb.Name
    } catch {
        # Unsaved / transient workbook (no FullName) — skip
    }
}

# Add-ins that are loaded (vbapm.xlam, etc.)
foreach ($addin in $Excel.AddIns) {
    try {
        if ($addin.Installed -or $addin.IsOpen) {
            $path = [System.IO.Path]::GetFullPath($addin.FullName)
            $OpenPaths[$path.ToLowerInvariant()] = $addin.Name
        }
    } catch { }
}

# ---------------------------------------------------------------
# Step 3 — iterate VBE projects and classify each one
# ---------------------------------------------------------------
$Ghosts   = @()
$Healthy  = @()

# Helper: test whether we can still interact with the parent workbook
# through COM.  Ghost projects often have lingering VBComponents but a
# disconnected / dead parent object.
function Test-ProjectResponsive {
    param($Project)
    $indicators = @{}

    # 3a. FileName — empty or inaccessible is a strong signal
    try {
        $indicators.FileName = $Project.FileName
    } catch {
        $indicators.FileName = $null
        $indicators.FileNameError = $_.Exception.Message
    }

    # 3b. VBComponents count — still accessible even for ghosts, but 0 is
    #     suspicious for a document project
    try {
        $indicators.ComponentCount = $Project.VBComponents.Count
    } catch {
        $indicators.ComponentCount = $null
        $indicators.ComponentError = $_.Exception.Message
    }

    # 3c. Protection — a locked project that throws may be a ghost
    try {
        $indicators.Protection = $Project.Protection
    } catch {
        $indicators.Protection = $null
        $indicators.ProtectionError = $_.Exception.Message
    }

    return $indicators
}

# Helper: determine whether this project is the vbapm add-in itself
function Test-IsVbapmAddin {
    param($Project)
    try {
        $fn = $Project.FileName
        return ($fn -and ($fn -like "*\vbapm.xlam" -or $fn -like "*/vbapm.xlam"))
    } catch {
        return $false
    }
}

foreach ($proj in $Excel.VBE.VBProjects) {
    $projName = ""
    try { $projName = $proj.Name } catch { $projName = "(unknown)" }

    # Filter by name when -Name is provided
    if ($Name -and $projName -ne $Name) { continue }

    # Skip the vbapm add-in's own project
    if (Test-IsVbapmAddin $proj) {
        if (-not $Quiet) { Write-Host "SKIP  $projName  [vbapm add-in]" -ForegroundColor DarkGray }
        continue
    }

    $indicators = Test-ProjectResponsive $proj

    # ---- Classification logic ----------------------------------
    $isGhost = $false
    $reason  = ""

    if (-not $indicators.FileName) {
        # No FileName at all — cannot cross-reference; treat as ghost
        # only if component count is non-zero (a real ghost still has
        # cached components).
        if ($indicators.ComponentCount -and $indicators.ComponentCount -gt 0) {
            $isGhost = $true
            $reason  = "FileName is inaccessible ($($indicators.FileNameError)) but VBComponents.Count = $($indicators.ComponentCount)"
        }
    } else {
        $normalized = [System.IO.Path]::GetFullPath($indicators.FileName).ToLowerInvariant()
        if (-not $OpenPaths.ContainsKey($normalized)) {
            $isGhost = $true
            $reason  = "File `"$($indicators.FileName)`" is not open in Workbooks or AddIns"
        }
    }

    # ---- Reporting ----------------------------------------------
    $entry = [PSCustomObject]@{
        Name            = $projName
        File            = $indicators.FileName
        Protection      = $indicators.Protection
        ComponentCount  = $indicators.ComponentCount
        IsGhost         = $isGhost
        Reason          = $reason
        FileNameError   = $indicators.FileNameError
        ComponentError  = $indicators.ComponentError
        ProtectionError = $indicators.ProtectionError
    }

    if ($isGhost) {
        $Ghosts += $entry
        if (-not $Quiet) {
            Write-Host "GHOST $projName" -ForegroundColor Red -NoNewline
            Write-Host "  | $reason"
        }
    } else {
        $Healthy += $entry
        if (-not $Quiet) {
            Write-Host "OK    $projName" -ForegroundColor Green -NoNewline
            Write-Host "  | matched: $($OpenPaths[$normalized])"
        }
    }
}

# ---------------------------------------------------------------
# Step 4 — output
# ---------------------------------------------------------------
if ($Raw) {
    # Single JSON array for machine consumption
    $all = @($Healthy) + @($Ghosts)
    ConvertTo-Json $all -Depth 3
} elseif ($Quiet) {
    # Compact summary
    $summary = [PSCustomObject]@{
        Total   = $Healthy.Count + $Ghosts.Count
        Healthy = $Healthy.Count
        Ghosts  = $Ghosts.Count
        GhostNames = @($Ghosts | ForEach-Object { $_.Name })
    }
    ConvertTo-Json $summary -Depth 2
} else {
    Write-Host ""
    Write-Host "===== SUMMARY =====" -ForegroundColor Cyan
    Write-Host "Healthy : $($Healthy.Count)" -ForegroundColor Green
    Write-Host "Ghosts  : $($Ghosts.Count)" -ForegroundColor Red
    if ($Ghosts.Count -gt 0) {
        Write-Host "Ghost project names: $($Ghosts.Name -join ', ')" -ForegroundColor Red
    }
}

exit 0
