# Test-WordComRotSelection.ps1
#
# Word-specific COM test. Word.Application.Hwnd is not populated reliably, so
# this script uses the active document window handle to identify each instance.

[CmdletBinding()]
param(
    [ValidateRange(2, 20)]
    [int]$Count = 3,
    [ValidateRange(0, 100)]
    [int]$LookupCount = 0,
    [ValidateRange(0, 10000)]
    [int]$DelayMilliseconds = 500,
    [switch]$ReleaseLookupReferences,
    [switch]$KeepLookupReferences
)

$ErrorActionPreference = 'Stop'

if ($LookupCount -eq 0) {
    $LookupCount = $Count
}

if ($PSVersionTable.PSEdition -eq 'Core') {
    $windowsPowerShell = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if (-not $windowsPowerShell) {
        throw 'This test requires Windows PowerShell for Word COM automation.'
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
    if ($KeepLookupReferences) {
        $forwardedArguments += '-KeepLookupReferences'
    }
    if ($ReleaseLookupReferences) {
        $forwardedArguments += '-ReleaseLookupReferences'
    }
    & $windowsPowerShell.Source $forwardedArguments
    exit $LASTEXITCODE
}

if (-not ('WordComSelectionNative' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class WordComSelectionNative {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(
        IntPtr windowHandle,
        out uint processId);
}
"@
}

function Get-WordIdentity {
    param([Parameter(Mandatory)][object]$WordApp)

    $hwnd = [int64]$WordApp.ActiveWindow.Hwnd
    [uint32]$processId = 0
    if ($hwnd -gt 0) {
        [WordComSelectionNative]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$processId) | Out-Null
    }

    [pscustomobject]@{
        Pid  = [int]$processId
        Hwnd = $hwnd
    }
}

function Remove-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject) {
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject) | Out-Null } catch {}
    }
}

function Get-WordDocuments {
    param([Parameter(Mandatory)][object]$WordApp)

    $documents = @()
    foreach ($document in $WordApp.Documents) {
        try {
            $documents += [string]$document.Name
        } finally {
            Remove-ComObject -ComObject $document
        }
    }
    return $documents
}

$wordType = [System.Type]::GetTypeFromProgID('Word.Application')
if ($null -eq $wordType) {
    throw 'Word.Application is not registered on this machine.'
}

$created = @()
$lookupObjects = @()
$lookupPids = @()
$lookupFailures = 0

try {
    Write-Output "Creating $Count Word instance(s) with one document each..."

    for ($index = 1; $index -le $Count; $index++) {
        $word = [System.Activator]::CreateInstance($wordType)
        $word.Visible = $true
        $document = $word.Documents.Add()
        $document.Saved = $true
        Remove-ComObject -ComObject $document

        $identity = Get-WordIdentity -WordApp $word
        $created += [pscustomobject]@{
            Order = $index
            Pid   = $identity.Pid
            Hwnd  = $identity.Hwnd
            App   = $word
        }

        Write-Output ("  created order={0} pid={1} hwnd={2}" -f $index, $identity.Pid, $identity.Hwnd)
        if ($DelayMilliseconds -gt 0) {
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }

    Write-Output ''
    Write-Output 'Marshal.GetActiveObject observations:'
    for ($index = 1; $index -le $LookupCount; $index++) {
        $activeWord = $null
        try {
            $activeWord = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
            $identity = Get-WordIdentity -WordApp $activeWord
            $documents = @(Get-WordDocuments -WordApp $activeWord)
            $lookupPids += $identity.Pid
            Write-Output ("  lookup={0} pid={1} hwnd={2} documents={3}" -f $index, $identity.Pid, $identity.Hwnd, $documents.Count)
            foreach ($documentName in $documents) {
                Write-Output "    - $documentName"
            }

            if ($ReleaseLookupReferences) {
                Remove-ComObject -ComObject $activeWord
            } else {
                $lookupObjects += $activeWord
            }
        } catch {
            $lookupFailures++
            Write-Output ("  lookup={0} failed: {1}" -f $index, $_.Exception.Message)
            Remove-ComObject -ComObject $activeWord
        }

        if ($DelayMilliseconds -gt 0) {
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }

    $createdPids = @($created | Where-Object { $_.Pid -gt 0 } | ForEach-Object { $_.Pid })
    Write-Output ''
    if ($createdPids.Count -lt 2 -or $lookupPids.Count -eq 0 -or @($lookupPids | Where-Object { $_ -le 0 }).Count -gt 0) {
        Write-Output 'RESULT: Inconclusive. Word did not expose enough window/PID information.'
    } elseif ($lookupFailures -gt 0) {
        Write-Output "RESULT: $lookupFailures lookup(s) failed after $($lookupPids.Count) successful lookup(s)."
    } elseif (($lookupPids | Select-Object -Unique).Count -eq 1) {
        Write-Output 'RESULT: Every lookup returned the same Word instance.'
    } else {
        Write-Output 'RESULT: Successive lookups returned different Word instances.'
    }
} finally {
    foreach ($lookupObject in $lookupObjects) {
        Remove-ComObject -ComObject $lookupObject
    }

    foreach ($createdInstance in $created) {
        try { $createdInstance.App.Quit() } catch {}
        Remove-ComObject -ComObject $createdInstance.App
    }
}

Write-Output 'Created Word instances were closed.'
