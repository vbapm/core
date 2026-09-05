<#
.SYNOPSIS
    Demonstrates that VB.NET GetObject can create a new Excel instance instead
    of attaching to an existing one.


.DESCRIPTION
    Records the Excel instances listed in instances.json and the Excel
    processes running before the call. It then calls
    Microsoft.VisualBasic.Interaction.GetObject with a null path and the
    Excel.Application class name. The result shows whether the returned
    process existed before the call.

.EXAMPLE
    .\scripts\ps\demo-vb-getobject.ps1
#>

$ErrorActionPreference = 'Stop'

$registryPath = Join-Path $env:TEMP 'Excel-Instances\instances.json'
$registryPids = @()
if (Test-Path -LiteralPath $registryPath) {
    $snapshot = Get-Content -LiteralPath $registryPath -Raw | ConvertFrom-Json
    $registryPids = @($snapshot.active | ForEach-Object { [int]$_.pid })
}

$beforeProcesses = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
$beforePids = @($beforeProcesses | ForEach-Object { [int]$_.Id })

Write-Output 'Excel instances before VB.NET GetObject:'
if ($registryPids.Count -eq 0) {
    Write-Output '  instances.json active PIDs: none'
} else {
    Write-Output "  instances.json active PIDs: $($registryPids -join ', ')"
}
if ($beforePids.Count -eq 0) {
    Write-Output '  running EXCEL.EXE PIDs: none'
} else {
    Write-Output "  running EXCEL.EXE PIDs: $($beforePids -join ', ')"
}

if (-not ('WindowProcess' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class WindowProcess {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(
        IntPtr windowHandle,
        out uint processId);
}
"@
}

try {
    $excel = [Microsoft.VisualBasic.Interaction]::GetObject($null, 'Excel.Application')
} catch {
    Write-Error "VB.GetObject could not find a running Excel instance: $($_.Exception.Message)"
    exit 1
}

try {
    [uint32]$processId = 0
    [WindowProcess]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$processId) | Out-Null

    $wasRunningBefore = $beforePids -contains $processId
    $wasInRegistryBefore = $registryPids -contains $processId

    Write-Output ''
    Write-Output 'VB.NET GetObject returned:'
    Write-Output "  ProcessId: $processId"
    Write-Output "  Hwnd:      $($excel.Hwnd)"
    Write-Output "  Visible:   $($excel.Visible)"
    Write-Output "  Workbooks: $($excel.Workbooks.Count)"
    Write-Output "  Was running before call: $wasRunningBefore"
    Write-Output "  Was listed in instances.json: $wasInRegistryBefore"

    if (-not $wasRunningBefore) {
        Write-Output ''
        Write-Output "RESULT: GetObject returned a new Excel process. PID $processId was not running before the call."
    } else {
        Write-Output ''
        Write-Output "RESULT: GetObject returned an Excel process that was already running. PID $processId was present before the call."
    }

    foreach ($workbook in $excel.Workbooks) {
        Write-Output "    - $($workbook.Name) [$($workbook.FullName)]"
    }
} finally {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
