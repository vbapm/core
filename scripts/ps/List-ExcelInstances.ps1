# List-ExcelInstances.ps1
#
# Sync the Excel instance registry, then list each active instance and its
# workbooks from instances.json.

[CmdletBinding()]
param(
    [Alias('Window')]
    [switch]$Windows,

    [string]$RegistryPath,

    [switch]$OutputJson
)

$ErrorActionPreference = 'Stop'

if ($env:EVOLUTION_JSON -eq '1' -and $env:SNAPSHOT_FILE) {
    $env:VBA_JSON_BASE64 = '1'
    $RegistryPath = $env:SNAPSHOT_FILE
}

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

if ($RegistryPath) {
    if (-not (Test-Path -LiteralPath $RegistryPath -PathType Leaf)) {
        throw "Registry snapshot not found: $RegistryPath"
    }

    $snapshot = Get-Content -LiteralPath $RegistryPath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
    $instances = @($snapshot.active)
    $inactive = @($snapshot.inactive)
} else {
    $syncScript = Join-Path $PSScriptRoot 'Sync-ExcelInstances.ps1'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript -PruneDead | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Could not sync the Excel instance registry. Exit code: $LASTEXITCODE"
    }

    $instances = @(Read-ExcelInstances)
    $inactive = @(Read-InactiveExcelInstances)
}

if ($OutputJson -or ($RegistryPath -and [string]$env:VBA_JSON_BASE64 -match '^(1|true|yes)$')) {
    $json = [ordered]@{
        active   = $instances
        inactive = $inactive
    } | ConvertTo-Json -Depth 20
    if ([string]$env:VBA_JSON_BASE64 -match '^(1|true|yes)$') {
        [Convert]::ToBase64String([System.Text.UTF8Encoding]::new($false).GetBytes($json))
    } else {
        [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
        $json
    }
    exit 0
}

$excelIcon = [char]::ConvertFromUtf32(0x1F4CA) # 📊 Excel icon
$visibleIcon = [char]::ConvertFromUtf32(0x1F441) + [char]0xFE0F # 👁️ Visible icon
$invisibleIcon = [char]::ConvertFromUtf32(0x1F648) # 🙈 Invisible icon
$workbookIcon = [char]::ConvertFromUtf32(0x1F4D6) # 📖 Workbook icon
$addinIcon = [char]::ConvertFromUtf32(0x1F4DA) # 📚 Add-in icon
$editableIcon = [char]0x270F + [char]0xFE0F # ✏️ Editable icon
$readOnlyIcon = [char]::ConvertFromUtf32(0x1F52D) # 🔭 Read-only icon
$ghostIcon = [char]::ConvertFromUtf32(0x1F47B) # 👻 Ghost workbook icon
$emptyIcon = [char]::ConvertFromUtf32(0x1F4ED) # 📭 Empty icon
$windowIcon = [char]::ConvertFromUtf32(0x1FA9F) # 🪟 Window icon
$comIcon = [char]::ConvertFromUtf32(0x2B50) # ⭐ COM-reachable icon
$windowScript = Join-Path $PSScriptRoot 'List-ExcelProcessWindows.ps1'
$verboseOutput = $VerbosePreference -eq 'Continue'

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
    Write-Host "$($instance.number)) $stateIcon$reachabilityIcon PID=$($instance.pid) visible=$($instance.visible) comReachable=$($instance.comReachable) owner=$($instance.owner) reason=$($instance.reason)" -ForegroundColor $stateColor

    $workbooks = @($instance.workbooks)
    if ($instance.comReachable) {
        if ($workbooks.Count -eq 0 -or $null -eq $workbooks[0]) {
            Write-Host "  $emptyIcon workbooks: none" -ForegroundColor DarkGray
        } else {
            $ghostByPath = @{}
            foreach ($ghost in @($instance.ghostWorkbooks)) {
                if ($ghost.file) {
                    try {
                        $ghostPath = [System.IO.Path]::GetFullPath([string]$ghost.file).ToLowerInvariant()
                        $ghostByPath[$ghostPath] = $ghost
                    } catch {}
                }
            }

            foreach ($workbook in $workbooks) {
                $isGhost = $false
                try {
                    if ($workbook -notmatch '^\[.*\]$') {
                        $workbookPath = [System.IO.Path]::GetFullPath([string]$workbook).ToLowerInvariant()
                        $isGhost = $ghostByPath.ContainsKey($workbookPath)
                    }
                } catch {}

                if ($isGhost) {
                    Write-Host "  $ghostIcon $workbook" -ForegroundColor Red
                } else {
                    $workbookKey = ([string]$workbook).ToLowerInvariant()
                    $workbookReadOnly = $false
                    if ($instance.workbookReadOnly -and $instance.workbookReadOnly.PSObject.Properties.Name -contains $workbookKey) {
                        $workbookReadOnly = [bool]$instance.workbookReadOnly.PSObject.Properties[$workbookKey].Value
                    }
                    $fileStateIcon = if ($workbookReadOnly) { $readOnlyIcon } else { $editableIcon }
                    Write-Host "  $workbookIcon$fileStateIcon $workbook" -ForegroundColor White
                }
            }
        }
    }

    if (-not $instance.comReachable -or $Windows) {
        if (Test-Path -LiteralPath $windowScript) {
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
    }

    $addins = @($instance.addins)
    $openAddins = @($addins | Where-Object { $_.isOpen })
    if ($openAddins.Count -gt 0) {
        foreach ($addin in $openAddins) {
            $addinPath = if ($addin.fullPath) { $addin.fullPath } else { $addin.name }
            $fileStateIcon = if ($addin.readOnly) { $readOnlyIcon } else { $editableIcon }
            Write-Host "  $addinIcon$fileStateIcon $addinPath" -ForegroundColor DarkCyan
        }
    }

    Write-Host ''
}
