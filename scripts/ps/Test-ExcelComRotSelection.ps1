# Test-ExcelComRotSelection.ps1
#
# Create several Excel COM instances in a known order, then compare the
# instance returned by Marshal.GetActiveObject with the creation order.

[CmdletBinding()]
param(
    [ValidateRange(2, 20)]
    [int]$Count = 3,
    [ValidateRange(0, 100)]
    [int]$LookupCount = 0,
    [ValidateRange(0, 10000)]
    [int]$DelayMilliseconds = 500,
    [switch]$ReleaseLookupReferences,
    [switch]$Visible,
    [switch]$KeepOpen
)

$ErrorActionPreference = 'Stop'

if ($LookupCount -eq 0) {
    $LookupCount = $Count
}

if ($PSVersionTable.PSEdition -eq 'Core') {
    $windowsPowerShell = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if (-not $windowsPowerShell) {
        throw 'This test requires Windows PowerShell because Marshal.GetActiveObject is unavailable in PowerShell Core.'
    }

    $forwardedArguments = @(
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        $PSCommandPath,
        '-Count',
        $Count,
        '-LookupCount',
        $LookupCount,
        '-DelayMilliseconds',
        $DelayMilliseconds
    )
    if ($ReleaseLookupReferences) {
        $forwardedArguments += '-ReleaseLookupReferences'
    }
    if ($Visible) {
        $forwardedArguments += '-Visible'
    }
    if ($KeepOpen) {
        $forwardedArguments += '-KeepOpen'
    }
    & $windowsPowerShell.Source $forwardedArguments
    exit $LASTEXITCODE
}

if (-not ('ComRotSelectionNative' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class ComRotSelectionNative {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(
        IntPtr windowHandle,
        out uint processId);
}
"@
}

function Get-ComIdentity {
    param([Parameter(Mandatory)][object]$ComObject)

    $hwnd = 0L
    try { $hwnd = [int64]$ComObject.Hwnd } catch {}

    $processId = 0
    if ($hwnd -gt 0) {
        [uint32]$nativeProcessId = 0
        [ComRotSelectionNative]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$nativeProcessId) | Out-Null
        $processId = [int]$nativeProcessId
    }

    [pscustomobject]@{
        Pid  = $processId
        Hwnd = $hwnd
        Type = $ComObject.GetType().FullName
    }
}

function Remove-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject) {
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject) | Out-Null } catch {}
    }
}

$excelType = [System.Type]::GetTypeFromProgID('Excel.Application')
if ($null -eq $excelType) {
    throw 'Excel.Application is not registered on this machine.'
}

$created = @()
$lookups = @()
$lookupObjects = @()
$lookupFailures = 0
$baselineIdentity = $null

try {
    try {
        $baselineObject = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        $baselineIdentity = Get-ComIdentity -ComObject $baselineObject
        Write-Output ("Pre-existing GetActiveObject result: pid={0} hwnd={1}" -f $baselineIdentity.Pid, $baselineIdentity.Hwnd)
        Remove-ComObject -ComObject $baselineObject
    } catch {
        Write-Output 'Pre-existing GetActiveObject result: none'
    }

    Write-Output "Creating $Count Excel instance(s) in order..."

    for ($index = 1; $index -le $Count; $index++) {
        $comObject = [System.Activator]::CreateInstance($excelType)
        try { $comObject.Visible = [bool]$Visible } catch {}

        $identity = Get-ComIdentity -ComObject $comObject
        $created += [pscustomobject]@{
            Order = $index
            Pid   = $identity.Pid
            Hwnd  = $identity.Hwnd
            Type  = $identity.Type
            App   = $comObject
        }

        Write-Output ("  created order={0} pid={1} hwnd={2}" -f $index, $identity.Pid, $identity.Hwnd)
        if ($DelayMilliseconds -gt 0) {
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }

    Write-Output ''
    Write-Output 'Marshal.GetActiveObject observations:'
    for ($index = 1; $index -le $LookupCount; $index++) {
        try {
            $activeObject = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
            $identity = Get-ComIdentity -ComObject $activeObject
            $lookups += $identity
            Write-Output ("  lookup={0} pid={1} hwnd={2}" -f $index, $identity.Pid, $identity.Hwnd)
            if ($ReleaseLookupReferences) {
                Remove-ComObject -ComObject $activeObject
            } else {
                $lookupObjects += $activeObject
            }
        } catch {
            $lookupFailures++
            Write-Output ("  lookup={0} failed: {1}" -f $index, $_.Exception.Message)
        }

        if ($DelayMilliseconds -gt 0) {
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }

    $createdPids = @($created | Where-Object { $_.Pid -gt 0 } | ForEach-Object { $_.Pid })
    $lookupPids = @($lookups | Where-Object { $_.Pid -gt 0 } | ForEach-Object { $_.Pid })

    Write-Output ''
    if ($createdPids.Count -lt 2 -or $lookupPids.Count -eq 0) {
        Write-Output 'RESULT: Inconclusive. Excel did not expose enough PID information for comparison.'
    } elseif ($lookupFailures -gt 0) {
        Write-Output "RESULT: $lookupFailures lookup(s) failed after $($lookupPids.Count) successful lookup(s)."
    } elseif (($lookupPids | Select-Object -Unique).Count -ne 1) {
        if ($lookupPids.Count -gt $createdPids.Count -and $lookupPids[$createdPids.Count] -eq $lookupPids[0]) {
            Write-Output 'RESULT: The lookup sequence wrapped back to the first instance.'
        } else {
            Write-Output 'RESULT: The returned instance changed between lookups.'
        }
    } elseif ($lookupPids[0] -eq $createdPids[0]) {
        Write-Output 'RESULT: This run supports oldest-instance selection.'
    } elseif ($lookupPids[0] -eq $createdPids[-1]) {
        Write-Output 'RESULT: This run contradicts oldest-instance selection; the newest instance was returned.'
    } elseif ($createdPids -contains $lookupPids[0]) {
        Write-Output 'RESULT: This run returned a created instance, but neither the oldest nor newest.'
    } elseif ($null -ne $baselineIdentity -and $lookupPids[0] -eq $baselineIdentity.Pid) {
        Write-Output 'RESULT: Inconclusive. GetActiveObject kept returning a pre-existing instance; close all existing instances and rerun.'
    } else {
        Write-Output 'RESULT: GetActiveObject returned a PID not found among the created instances.'
    }
} finally {
    if (-not $KeepOpen) {
        foreach ($createdInstance in $created) {
            try { $createdInstance.App.Quit() } catch {}
        }
    }

    foreach ($lookupObject in $lookupObjects) {
        Remove-ComObject -ComObject $lookupObject
    }

    foreach ($createdInstance in $created) {
        Remove-ComObject -ComObject $createdInstance.App
    }
}

if ($KeepOpen) {
    Write-Output 'Created instances were left open because -KeepOpen was specified.'
} else {
    Write-Output 'Created instances were closed.'
}
