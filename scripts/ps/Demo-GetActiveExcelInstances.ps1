# Demo-GetActiveExcelInstances.ps1
#
# Experiment: on Excel installations where successive GetActiveObject calls
# expose different registrations, loop through those instances and list their
# workbooks. This is observed Excel behavior, not a general COM contract.

[CmdletBinding()]
param(
    [ValidateRange(1, 100)]
    [int]$MaximumInstances = 100
)

$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSEdition -eq 'Core') {
    $windowsPowerShell = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if (-not $windowsPowerShell) {
        throw 'This demo requires Windows PowerShell because Marshal.GetActiveObject is unavailable in PowerShell Core.'
    }

    & $windowsPowerShell.Source -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath -MaximumInstances $MaximumInstances
    exit $LASTEXITCODE
}

if (-not ('ActiveExcelInstancesNative' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class ActiveExcelInstancesNative {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(
        IntPtr windowHandle,
        out uint processId);
}
"@
}

function Get-ExcelProcessId {
    param([Parameter(Mandatory)][object]$ExcelApp)

    [uint32]$processId = 0
    $hwnd = [int64]$ExcelApp.Hwnd
    if ($hwnd -gt 0) {
        [ActiveExcelInstancesNative]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$processId) | Out-Null
    }
    return [int]$processId
}

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject) {
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject) | Out-Null } catch {}
    }
}

$excelIcon = [char]::ConvertFromUtf32(0x1F4CA)
$workbookIcon = [char]::ConvertFromUtf32(0x1F4D6)
$emptyIcon = [char]::ConvertFromUtf32(0x1F4ED)
$seen = @{}
$count = 0

Write-Host "$excelIcon Successive GetActiveObject Excel instances" -ForegroundColor Cyan

while ($count -lt $MaximumInstances) {
    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        $processId = Get-ExcelProcessId -ExcelApp $excel
        $hwnd = [int64]$excel.Hwnd
        $identity = "$processId/$hwnd"

        if ($seen.ContainsKey($identity)) {
            Write-Host "Repeated instance PID=$processId Hwnd=$hwnd. Stopping." -ForegroundColor Yellow
            break
        }
        $seen[$identity] = $true
        $count++

        $workbooks = @()
        foreach ($workbook in $excel.Workbooks) {
            try {
                $fullName = [string]$workbook.FullName
                $name = [string]$workbook.Name
                $workbooks += if ($fullName -and $fullName -ne $name) {
                    $fullName -replace '\\', '/'
                } else {
                    "[$name]"
                }
            } finally {
                Release-ComObject -ComObject $workbook
            }
        }

        Write-Host "Instance $count" -ForegroundColor Green
        Write-Host "  PID:       $processId"
        Write-Host "  Hwnd:      $hwnd"
        Write-Host "  Visible:   $([bool]$excel.Visible)"
        Write-Host "  Workbooks: $($workbooks.Count)"
        if ($workbooks.Count -eq 0) {
            Write-Host "  $emptyIcon none" -ForegroundColor DarkGray
        } else {
            foreach ($workbookName in $workbooks) {
                Write-Host "  $workbookIcon $workbookName" -ForegroundColor White
            }
        }
        Write-Host ''
    } catch {
        Write-Host 'No more COM-reachable Excel instances.' -ForegroundColor DarkGray
        break
    } finally {
        Release-ComObject -ComObject $excel
    }
}

if ($count -eq $MaximumInstances) {
    Write-Host "Stopped after the safety limit of $MaximumInstances instances." -ForegroundColor Yellow
}
