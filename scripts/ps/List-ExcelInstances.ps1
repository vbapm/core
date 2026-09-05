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

$excelIcon = [char]::ConvertFromUtf32(0x1F4CA) # 📊 Excel icon
$visibleIcon = [char]::ConvertFromUtf32(0x1F441) + [char]0xFE0F # 👁️ Visible icon
$invisibleIcon = [char]::ConvertFromUtf32(0x1F648) # 🙈 Invisible icon
$workbookIcon = [char]::ConvertFromUtf32(0x1F4D6) # 📖 Workbook icon
$emptyIcon = [char]::ConvertFromUtf32(0x1F4ED) # 📭 Empty icon
$windowIcon = [char]::ConvertFromUtf32(0x1FA9F) # 🪟 Window icon
$comIcon = [char]::ConvertFromUtf32(0x2B50) # ⭐ COM-reachable icon
$windowScript = Join-Path $PSScriptRoot 'List-ExcelProcessWindows.ps1'
$verboseOutput = $VerbosePreference -eq 'Continue'

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

    $reachabilityIcon = if ($instance.comReachable) { $comIcon } else { '' }
    Write-Host "$stateIcon$reachabilityIcon PID=$($instance.pid) visible=$($instance.visible) comReachable=$($instance.comReachable) owner=$($instance.owner) reason=$($instance.reason)" -ForegroundColor $stateColor

    $workbooks = @($instance.workbooks)
    if ($instance.comReachable) {
        if ($workbooks.Count -eq 0 -or $null -eq $workbooks[0]) {
            Write-Host "  $emptyIcon workbooks: none" -ForegroundColor DarkGray
        } else {
            foreach ($workbook in $workbooks) {
                Write-Host "  $workbookIcon $workbook" -ForegroundColor White
            }
        }
    } elseif (Test-Path -LiteralPath $windowScript) {
        $windowLines = @(
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $windowScript -ProcessId ([int]$instance.pid) |
                Where-Object { $_ -match 'class=EXCEL7' -and $_ -match 'title=' }
        )

        if ($windowLines.Count -eq 0) {
            Write-Host "  $emptyIcon EXCEL7 window titles: none" -ForegroundColor DarkGray
        } else {
            foreach ($windowLine in $windowLines) {
                if ($verboseOutput) {
                    if ($windowLine -match 'title=(.*)$') {
                        $windowTitle = $matches[1]
                        $windowDetails = $windowLine.Substring(0, $windowLine.Length - $matches[0].Length).TrimEnd()
                        Write-Host "  $windowIcon  $windowTitle [$windowDetails]" -ForegroundColor White
                    } else {
                        Write-Host "  $windowIcon  $($windowLine.Trim())" -ForegroundColor White
                    }
                } elseif ($windowLine -match 'title=(.*)$') {
                    Write-Host "  $windowIcon  $($matches[1])" -ForegroundColor White
                }
            }
        }
    } else {
        Write-Host "  $emptyIcon window titles: unavailable" -ForegroundColor DarkGray
    }

    Write-Host ''
}
