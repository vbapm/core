# Close-ActiveExcelnstance.ps1
#
# Close the Excel.Application object returned by Marshal.GetActiveObject.
# This targets the COM-reachable instance, not an arbitrary EXCEL.EXE PID.

[CmdletBinding()]
param(
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSEdition -eq 'Core') {
    $windowsPowerShell = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if (-not $windowsPowerShell) {
        Write-Error 'This script requires Windows PowerShell for Marshal.GetActiveObject.'
        exit 1
    }

    $forwardedArguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
    if ($WhatIf) {
        $forwardedArguments += '-WhatIf'
    }
    & $windowsPowerShell.Source $forwardedArguments
    exit $LASTEXITCODE
}

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

$excel = $null
try {
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    } catch {
        Write-Output 'No COM-reachable Excel instance found.'
        exit 0
    }

    $processId = 0
    try {
        $processId = Get-ExcelProcessId -ExcelApp $excel
    } catch {
        # The COM object can still be closed even if its PID cannot be resolved.
    }

    if ($WhatIf) {
        if ($processId) {
            Write-Output "[WhatIf] would close COM-reachable Excel instance pid=$processId"
        } else {
            Write-Output '[WhatIf] would close the COM-reachable Excel instance.'
        }
        exit 0
    }

    $excel.Quit()
    if ($processId) {
        Write-Output "Closed COM-reachable Excel instance pid=$processId."
        Remove-ExcelInstance -InstanceId $processId -DeactivateReason 'closed-active'
    } else {
        Write-Output 'Closed COM-reachable Excel instance.'
    }
} finally {
    if ($null -ne $excel) {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
}

$syncScript = Join-Path $PSScriptRoot 'Sync-ExcelInstances.ps1'
if (Test-Path -LiteralPath $syncScript) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript -PruneDead
}
