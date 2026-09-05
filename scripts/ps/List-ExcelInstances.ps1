# List-ExcelInstances.ps1
#
# Sync the Excel instance registry, then list each active instance and its
# workbooks from instances.json.

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$syncScript = Join-Path $PSScriptRoot 'Sync-ExcelInstances.ps1'
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript -PruneDead | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Could not sync the Excel instance registry. Exit code: $LASTEXITCODE"
}

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

$excelIcon = [char]::ConvertFromUtf32(0x1F4CA)
$visibleIcon = [char]::ConvertFromUtf32(0x1F7E2)
$invisibleIcon = [char]::ConvertFromUtf32(0x1F534)
$workbookIcon = [char]::ConvertFromUtf32(0x1F4D6)
$emptyIcon = [char]::ConvertFromUtf32(0x1F4ED)

$instances = @(Read-ExcelInstances)
if ($instances.Count -eq 0) {
    Write-Host "$excelIcon No Excel instances found." -ForegroundColor Yellow
    exit 0
}

Write-Host "$excelIcon Excel instances: $($instances.Count)" -ForegroundColor Cyan
foreach ($instance in $instances) {
    if ($instance.visible) {
        $stateIcon = $visibleIcon
        $stateColor = 'Green'
    } else {
        $stateIcon = $invisibleIcon
        $stateColor = 'DarkYellow'
    }

    Write-Host "$stateIcon PID=$($instance.pid) visible=$($instance.visible) owner=$($instance.owner) reason=$($instance.reason)" -ForegroundColor $stateColor

    $workbooks = @($instance.workbooks)
    if ($workbooks.Count -eq 0 -or $null -eq $workbooks[0]) {
        Write-Host "  $emptyIcon workbooks: none" -ForegroundColor DarkGray
        continue
    }

    foreach ($workbook in $workbooks) {
        Write-Host "  $workbookIcon $workbook" -ForegroundColor White
    }

    Write-Host ''
}
