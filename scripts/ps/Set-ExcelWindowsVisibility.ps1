# Set-ExcelWindowsVisibility.ps1
#
# Shared ROT-based visibility implementation for one Excel.Application.

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, Position = 0)]
    [Alias('Pid')]
    [int]$ProcessId,

    [Parameter(Mandatory)]
    [bool]$Visible
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

$apps = @(Get-ExcelApplications)
$matchedApp = $null
try {
    foreach ($app in $apps) {
        try {
            if ((Get-ExcelProcessId -ExcelApp $app) -eq $ProcessId) {
                $matchedApp = $app
                break
            }
        } catch {}
    }

    if ($null -eq $matchedApp) {
        throw "No ROT-registered Excel.Application found for PID $ProcessId."
    }

    $action = if ($Visible) { 'Set Excel application and VBE visible' } else { 'Set Excel application and VBE invisible' }
    $target = "Excel PID $ProcessId Hwnd=$($matchedApp.Hwnd)"
    if ($PSCmdlet.ShouldProcess($target, $action)) {
        if ($Visible) {
            $matchedApp.Visible = $true
        }

        try {
            $matchedApp.VBE.MainWindow.Visible = $Visible
        } catch {
            Write-Warning "Could not set VBE visibility for Excel PID ${ProcessId}: $($_.Exception.Message)"
        }

        if (-not $Visible) {
            $matchedApp.Visible = $false
        }

        Write-Output "$action`: $target"
    } else {
        Write-Output "Would $($action.ToLowerInvariant()): $target"
    }
} finally {
    foreach ($app in $apps) {
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null } catch {}
    }
}
