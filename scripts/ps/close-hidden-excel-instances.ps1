<#
.SYNOPSIS
    Closes hidden / orphaned Excel instances that may be holding ghost
    VBE project references.  Visible instances are never touched.

.DESCRIPTION
    After vbapm build/export/update operations, hidden Excel instances
    can be left behind.  These abandoned processes hold COM references
    to workbooks that appear as "ghost" projects in the visible VBE.

    This script:
      1. Lists all Excel processes (visible + hidden)
      2. Attempts a graceful COM Quit on the ROT-registered instance
      3. Force-kills any remaining hidden instances

.PARAMETER Force
    Skip the confirmation prompt.

.PARAMETER WhatIf
    Preview which instances would be closed without actually doing it.

.PARAMETER TimeoutSeconds
    How long to wait for graceful COM shutdown before force-killing
    (default: 5 seconds).

.EXAMPLE
    .\scripts\ps\close-hidden-excel-instances.ps1

.EXAMPLE
    .\scripts\ps\close-hidden-excel-instances.ps1 -Force

.EXAMPLE
    .\scripts\ps\close-hidden-excel-instances.ps1 -WhatIf
#>

param(
    [switch]$Force,
    [switch]$WhatIf,
    [int]$TimeoutSeconds = 5
)

# ---------------------------------------------------------------
# PowerShell Core → re-invoke with Windows PowerShell for COM
# ---------------------------------------------------------------
if ($PSVersionTable.PSEdition -eq 'Core') {
    $winPS = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($winPS) {
        $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
        if ($Force)           { $psArgs += '-Force' }
        if ($WhatIf)          { $psArgs += '-WhatIf' }
        if ($TimeoutSeconds)  { $psArgs += '-TimeoutSeconds'; $psArgs += $TimeoutSeconds }
        & powershell.exe $psArgs
        exit $LASTEXITCODE
    }
}

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------
# Step 1 — enumerate Excel processes
# ---------------------------------------------------------------
$allProcs = @(Get-Process excel -ErrorAction SilentlyContinue)
if ($allProcs.Count -eq 0) {
    Write-Host "No Excel processes found (Task Manager should also show zero)." -ForegroundColor Green
    Write-Host "Nothing to close." -ForegroundColor DarkGray
    exit 0
}

$visibleProcs = @()
$hiddenProcs  = @()

foreach ($p in $allProcs) {
    $entry = [PSCustomObject]@{
        PID        = $p.Id
        StartTime  = $p.StartTime
        MemoryMB   = [math]::Round($p.WorkingSet64 / 1MB, 1)
        MainTitle  = $p.MainWindowTitle
        IsHidden   = [string]::IsNullOrEmpty($p.MainWindowTitle)
    }
    if ($entry.IsHidden) {
        $hiddenProcs += $entry
    } else {
        $visibleProcs += $entry
    }
}

Write-Host "===== Excel Instances =====" -ForegroundColor Cyan
Write-Host "Visible : $($visibleProcs.Count)"
foreach ($v in $visibleProcs) {
    Write-Host "  PID $($v.PID)  ""$($v.MainTitle)""  Started: $($v.StartTime)  Mem: $($v.MemoryMB)MB" -ForegroundColor Green
}
Write-Host "Hidden  : $($hiddenProcs.Count)"
foreach ($h in $hiddenProcs) {
    Write-Host "  PID $($h.PID)  Started: $($h.StartTime)  Mem: $($h.MemoryMB)MB" -ForegroundColor Yellow
}

if ($hiddenProcs.Count -eq 0) {
    Write-Host ""
    Write-Host "No hidden instances to close." -ForegroundColor Green
    exit 0
}

# ---------------------------------------------------------------
# Step 2 — try graceful COM shutdown on ROT-registered instance
# ---------------------------------------------------------------
Write-Host ""
Write-Host "===== Attempting graceful COM shutdown =====" -ForegroundColor Cyan

try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    if ($excel) {
        $rotPid = 0
        try { $rotPid = Get-Process -Id $excel.Hwnd -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id } catch { }
        
        Write-Host "ROT instance acquired. Workbooks open: $($excel.Workbooks.Count)"
        foreach ($wb in $excel.Workbooks) {
            Write-Host "  $($wb.Name)  Saved=$($wb.Saved)"
        }

        if (-not $WhatIf) {
            # Close workbooks then quit
            foreach ($wb in $excel.Workbooks) {
                try { $wb.Close($false) } catch { }
            }
            $excel.Quit()
            Write-Host "COM Quit() called. Waiting up to ${TimeoutSeconds}s..." -ForegroundColor Yellow

            $waited = 0
            while ($waited -lt $TimeoutSeconds) {
                Start-Sleep -Seconds 1
                $waited++
                $stillAlive = Get-Process excel -ErrorAction SilentlyContinue | Where-Object { $_.Id -in ($hiddenProcs.PID) }
                if (-not $stillAlive) {
                    Write-Host "All hidden instances exited gracefully." -ForegroundColor Green
                    break
                }
            }
        } else {
            Write-Host "[WhatIf] Would call Quit() on ROT instance" -ForegroundColor DarkGray
        }

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
} catch {
    Write-Host "Could not attach to ROT instance: $($_.Exception.Message)" -ForegroundColor DarkGray
}

# ---------------------------------------------------------------
# Step 3 — force-kill remaining hidden instances
# ---------------------------------------------------------------
$remaining = @(Get-Process excel -ErrorAction SilentlyContinue | Where-Object {
    [string]::IsNullOrEmpty($_.MainWindowTitle)
})

if ($remaining.Count -eq 0) {
    Write-Host ""
    Write-Host "No hidden instances remaining." -ForegroundColor Green
    exit 0
}

Write-Host ""
Write-Host "===== Force-killing $($remaining.Count) remaining hidden instance(s) =====" -ForegroundColor Cyan
foreach ($r in $remaining) {
    Write-Host "  PID $($r.Id)  Started: $($r.StartTime)" -ForegroundColor Red
}

if ($WhatIf) {
    Write-Host "[WhatIf] Would Stop-Process -Force on the above PIDs" -ForegroundColor DarkGray
    exit 0
}

if (-not $Force) {
    $confirm = Read-Host "`nKill these hidden Excel instances? [y/N]"
    if ($confirm -notmatch '^[yY]') {
        Write-Host "Aborted." -ForegroundColor Yellow
        exit 0
    }
}

foreach ($r in $remaining) {
    try {
        Stop-Process -Id $r.Id -Force
        Write-Host "  Killed PID $($r.Id)" -ForegroundColor Red
    } catch {
        Write-Host "  FAILED to kill PID $($r.Id): $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ---------------------------------------------------------------
# Step 4 — verify
# ---------------------------------------------------------------
Write-Host ""
$final = @(Get-Process excel -ErrorAction SilentlyContinue)
if ($final.Count -eq 0) {
    Write-Host "All Excel instances closed." -ForegroundColor Green
} else {
    $finalHidden = @($final | Where-Object { [string]::IsNullOrEmpty($_.MainWindowTitle) })
    Write-Host "Remaining: $($final.Count) total, $($finalHidden.Count) hidden"
}
