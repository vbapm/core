<#
.SYNOPSIS
    Creates a new Word instance as a separate COM process.

.DESCRIPTION
    Uses COM activation to create a new Word.Application object. Document
    operations use that returned object directly, so they cannot be redirected
    to another Word instance. Use -Invisible to hide the new instance.
    By default, creates a blank document.

.PARAMETER File
    Path to a Word document to open in the new instance.

.PARAMETER NoBlank
    Do not create a blank document.

.PARAMETER Invisible
    Start Word without showing its window.

.EXAMPLE
    .\scripts\ps\Create-NewWord.ps1

.EXAMPLE
    .\scripts\ps\Create-NewWord.ps1 -File "C:\my-project\document.docx"

.EXAMPLE
    .\scripts\ps\Create-NewWord.ps1 -NoBlank

.EXAMPLE
    .\scripts\ps\Create-NewWord.ps1 -Invisible
#>

param(
    [string]$File,
    [switch]$NoBlank,
    [switch]$Invisible
)

if ($File -eq '--invisible') {
    $File = $null
    $Invisible = $true
}

if ($PSVersionTable.PSEdition -eq 'Core') {
    $windowsPowerShell = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($windowsPowerShell) {
        $forwardedArguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
        if ($File) { $forwardedArguments += '-File'; $forwardedArguments += $File }
        if ($NoBlank) { $forwardedArguments += '-NoBlank' }
        if ($Invisible) { $forwardedArguments += '-Invisible' }
        & $windowsPowerShell.Source $forwardedArguments
        exit $LASTEXITCODE
    }
}

$ErrorActionPreference = 'Stop'

if (-not ('CreateNewWordWindow' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class CreateNewWordWindow {
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr windowHandle, int command);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(
        IntPtr windowHandle,
        out uint processId);
}
"@
}

function Get-CreatedWordProcessId {
    param([Parameter(Mandatory)][object]$WordApp)

    try {
        $hwnd = [int64]$WordApp.ActiveWindow.Hwnd
        [uint32]$processId = 0
        [CreateNewWordWindow]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$processId) | Out-Null
        if ($processId -gt 0) {
            return [int]$processId
        }
    } catch {}

    return $null
}

function Set-WordVisibility {
    param(
        [Parameter(Mandatory)][object]$WordApp,
        [bool]$Visible
    )

    $WordApp.Visible = $Visible
    if (-not $Visible) {
        try {
            $hwnd = [int64]$WordApp.ActiveWindow.Hwnd
            if ($hwnd -gt 0) {
                [CreateNewWordWindow]::ShowWindow([IntPtr]$hwnd, 0) | Out-Null
            }
        } catch {}
    }
}

function Remove-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject) {
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject) | Out-Null } catch {}
    }
}

Write-Host 'Launching new Word process...' -ForegroundColor Cyan

$resolved = $null
if ($File) {
    if (-not (Test-Path -LiteralPath $File)) {
        Write-Error "File not found: $File"
        exit 1
    }
    $resolved = (Resolve-Path -LiteralPath $File).Path
    Write-Host "Opening: $resolved" -ForegroundColor Yellow
}

$word = $null
$document = $null
try {
    $wordType = [System.Type]::GetTypeFromProgID('Word.Application')
    if ($null -eq $wordType) {
        throw 'Word.Application is not registered on this machine.'
    }

    $word = [System.Activator]::CreateInstance($wordType)
    $Word.Visible = -not $Invisible

    if ($resolved) {
        $document = $word.Documents.Open($resolved)
        Write-Host "Document opened: $resolved" -ForegroundColor Green
    } elseif (-not $NoBlank) {
        $document = $word.Documents.Add()
        Write-Host 'Blank document created.' -ForegroundColor Green
    }

    $processId = Get-CreatedWordProcessId -WordApp $word
    if ($processId) {
        Write-Host "New instance started: PID $processId" -ForegroundColor Green
    } else {
        Write-Host 'New instance started: PID unknown' -ForegroundColor Green
    }

    if ($Invisible) {
        Set-WordVisibility -WordApp $word -Visible $false
    }
} finally {
    Remove-ComObject -ComObject $document
    Remove-ComObject -ComObject $word
}

Write-Host ''
Write-Host 'The Word instance was left running.'
