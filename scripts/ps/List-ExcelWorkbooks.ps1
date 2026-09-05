# List-ExcelWorkbooks.ps1
#
# List each workbook exposed by every COM-reachable Excel.Application instance.
# Use -Json when another script needs structured output.

[CmdletBinding()]
param(
    [switch]$Json
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

$workbookRecords = @()
$instanceNumber = 0

foreach ($excel in @(Get-ExcelApplications)) {
    $instanceNumber++

    try {
        $processId = Get-ExcelProcessId -ExcelApp $excel
        $visible = [bool]$excel.Visible

        foreach ($workbook in @(Get-ExcelWorkbooks -ExcelApp $excel)) {
            if ([string]::IsNullOrWhiteSpace([string]$workbook)) {
                continue
            }

            $workbookRecords += [pscustomobject]@{
                Instance  = $instanceNumber
                ProcessId = $processId
                Visible   = $visible
                Workbook  = [string]$workbook
            }
        }
    } catch {
        Write-Warning "Could not read Excel instance ${instanceNumber}: $($_.Exception.Message)"
    } finally {
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null } catch {}
    }
}

if ($Json) {
    if ($workbookRecords.Count -eq 0) {
        Write-Output '[]'
    } else {
        $workbookRecords | ConvertTo-Json -Depth 3
    }
    exit 0
}

if ($workbookRecords.Count -eq 0) {
    Write-Output 'No open workbooks found in COM-reachable Excel instances.'
    exit 0
}

Write-Output "Open workbooks: $($workbookRecords.Count)"
foreach ($record in $workbookRecords) {
    Write-Output ("PID={0} visible={1} workbook={2}" -f $record.ProcessId, $record.Visible, $record.Workbook)
}
