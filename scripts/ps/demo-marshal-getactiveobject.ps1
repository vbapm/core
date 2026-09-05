<#
.SYNOPSIS
    Demonstrates that Marshal.GetActiveObject can return an existing Excel
    instance instead of creating a new one.

.DESCRIPTION
    Records the active Excel instances listed in instances.json and the Excel
    processes running before the call. It then calls
    System.Runtime.InteropServices.Marshal.GetActiveObject and proves that the
    returned process existed before the call. This API retrieves an existing
    COM object; it does not create Excel.

    PowerShell 7 does not expose Marshal.GetActiveObject in the .NET runtime
    used here, so a PowerShell Core invocation is forwarded to Windows
    PowerShell 5.1 automatically.

.EXAMPLE
    .\scripts\ps\demo-marshal-getactiveobject.ps1
#>

$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSEdition -eq 'Core') {
    $windowsPowerShell = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if (-not $windowsPowerShell) {
        Write-Error 'Marshal.GetActiveObject requires Windows PowerShell on this machine.'
        exit 1
    }

    & $windowsPowerShell.Source -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath @args
    exit $LASTEXITCODE
}

$registryPath = Join-Path $env:TEMP 'Excel-Instances\instances.json'
$registryPids = @()
if (Test-Path -LiteralPath $registryPath) {
    $snapshot = Get-Content -LiteralPath $registryPath -Raw | ConvertFrom-Json
    $registryPids = @($snapshot.active | ForEach-Object { [int]$_.pid })
}

$beforeProcesses = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
$beforePids = @($beforeProcesses | ForEach-Object { [int]$_.Id })

Write-Output 'Excel instances before Marshal.GetActiveObject:'
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

if (-not ('MarshalWindowProcess' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class MarshalWindowProcess {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(
        IntPtr windowHandle,
        out uint processId);
}
"@
}

try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
    Write-Host ([char]::ConvertFromUtf32(0x1F534)) -ForegroundColor Red -NoNewline
    Write-Host ' ' -NoNewline
    Write-Output 'No running Excel instance was found to attach to.'
    exit 0
}

try {
    [uint32]$processId = 0
    [MarshalWindowProcess]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$processId) | Out-Null

    $wasRunningBefore = $beforePids -contains $processId
    $wasInRegistryBefore = $registryPids -contains $processId

    Write-Output ''
    Write-Output 'Marshal.GetActiveObject returned:'
    Write-Output "  ProcessId: $processId"
    Write-Output "  Hwnd:      $($excel.Hwnd)"
    Write-Output "  Visible:   $($excel.Visible)"
    Write-Output "  Workbooks: $($excel.Workbooks.Count)"
    Write-Output "  Was running before call: $wasRunningBefore"
    Write-Output "  Was listed in instances.json: $wasInRegistryBefore"

    if ($wasRunningBefore) {
        Write-Output ''
        Write-Output "RESULT: Marshal.GetActiveObject returned an existing Excel process. PID $processId was present before the call."
    } else {
        Write-Output ''
        Write-Output "RESULT: The returned PID was not running before the call. PID $processId was not expected from an attach-only API."
    }

    foreach ($workbook in $excel.Workbooks) {
        Write-Output "    - $($workbook.Name) [$($workbook.FullName)]"
    }
} finally {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
