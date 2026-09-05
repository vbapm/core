# List-ExcelWindows.ps1
#
# List COM-reachable Excel applications and every workbook window they own.
# This distinguishes application/process identity from individual window identity.

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

$applications = @(Get-ExcelApplications)
if ($applications.Count -eq 0) {
    Write-Output 'No COM-reachable Excel instances found.'
    exit 0
}

Write-Output "COM-reachable Excel applications: $($applications.Count)"

$applicationIndex = 0
foreach ($excel in $applications) {
    $applicationIndex++
    try {
        $applicationHwnd = [int64]$excel.Hwnd
        $processId = Get-ExcelProcessId -ExcelApp $excel
        $windows = @($excel.Windows)

        Write-Output ''
        Write-Output "Application $applicationIndex"
        Write-Output "  PID:              $processId"
        Write-Output "  Application.Hwnd: $applicationHwnd"
        Write-Output "  Visible:          $([bool]$excel.Visible)"
        Write-Output "  Excel windows:    $($windows.Count)"

        if ($windows.Count -eq 0) {
            Write-Output '  No Excel workbook windows found.'
            continue
        }

        $windowIndex = 0
        foreach ($window in $windows) {
            $windowIndex++
            $workbookName = '<unknown>'
            try { $workbookName = [string]$window.Parent.Name } catch {}

            Write-Output "  Window $windowIndex"
            Write-Output "    Hwnd:     $([int64]$window.Hwnd)"
            Write-Output "    Caption:  $([string]$window.Caption)"
            Write-Output "    Workbook: $workbookName"

            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($window) | Out-Null
        }
    } finally {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
}
