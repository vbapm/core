# Demo-GetExcelApplications.ps1
#
# List the Excel.Application objects currently exposed through the COM Running
# Object Table and print the workbooks from each object.

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

$excelIcon = [char]::ConvertFromUtf32(0x1F4CA)
$workbookIcon = [char]::ConvertFromUtf32(0x1F4D6)
$emptyIcon = [char]::ConvertFromUtf32(0x1F4ED)

$applications = @(Get-ExcelApplications)
if ($applications.Count -eq 0) {
    Write-Host "$excelIcon No COM-reachable Excel instances found." -ForegroundColor Yellow
    exit 0
}

Write-Host "$excelIcon COM-reachable Excel instances: $($applications.Count)" -ForegroundColor Cyan

$index = 0
foreach ($excel in $applications) {
    $index++
    try {
        $processId = Get-ExcelProcessId -ExcelApp $excel
        $visible = [bool]$excel.Visible
        $workbooks = @(Get-ExcelWorkbooks -ExcelApp $excel)

        Write-Host "Excel instance $index" -ForegroundColor Green
        Write-Host "  PID:       $processId"
        Write-Host "  Hwnd:      $($excel.Hwnd)"
        Write-Host "  Visible:   $visible"
        Write-Host "  Workbooks: $($workbooks.Count)"

        if ($workbooks.Count -eq 0) {
            Write-Host "  $emptyIcon none" -ForegroundColor DarkGray
        } else {
            foreach ($workbook in $workbooks) {
                Write-Host "  $workbookIcon $workbook" -ForegroundColor White
            }
        }
    } finally {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }

    Write-Host ''
}
