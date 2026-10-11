# Ensure-ExcelInstancesClean.ps1
#
# Guard run immediately before the e2e integration suite. It terminates
# invisible rogue/untracked EXCEL.EXE processes and stale hidden instances
# owned by prior e2e sessions. Visible Excel processes are never killed,
# regardless of their registry owner; they are reported as warnings while the
# suite continues.
#
# The e2e suite itself creates *background* instances (VBA_BACKGROUND_BUILD=1)
# and registers each one, so in the common case there is nothing running when
# this guard fires.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\scripts\ps\Ensure-ExcelInstancesClean.ps1
#   powershell -ExecutionPolicy Bypass -File .\scripts\ps\Ensure-ExcelInstancesClean.ps1 -Clean

[CmdletBinding()]
param(
	[switch]$Clean
)

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

function Write-PreflightLog {
	param([string]$Message)

	$stamp = [DateTime]::UtcNow.ToString('o')
	Write-Output "[$stamp] $Message"
}

Write-PreflightLog "Starting Excel instance preflight (Clean=$Clean)."

if ($Clean) {
	$warmMetadata = @(Get-ChildItem -LiteralPath (Join-Path $env:TEMP 'Excel-Instances') -Filter 'warm-session*.json' -File -ErrorAction SilentlyContinue)
	Write-PreflightLog "Removing $($warmMetadata.Count) warm-session metadata file(s)."
	$warmMetadata | Remove-Item -Force -ErrorAction SilentlyContinue
}

Sync-ExcelInstanceSnapshot | Out-Null
$activeByPid = @{}
foreach ($entry in @(Read-ExcelInstances)) {
	$activeByPid[[int]$entry.pid] = $entry
}

$excelProcesses = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
Write-PreflightLog "Found $($excelProcesses.Count) Excel process(es) and $($activeByPid.Count) tracked active instance(s)."

$visibleProtected = @()
$invisibleCleanup = @()
foreach ($process in $excelProcesses) {
	$entry = $activeByPid[[int]$process.Id]
	$owner = if ($entry) { [string]$entry.owner } else { 'untracked' }
	$isRogue = -not $entry -or $owner -match '^(rogue|unknown|unknown-owner|untracked)$'
	$isCleanTracked = $Clean -and $null -ne $entry
	if (-not $isRogue -and -not $isCleanTracked) { continue }

	if ([bool]($process.MainWindowHandle -ne 0)) {
		$visibleProtected += [pscustomobject]@{ Process = $process; CleanTracked = $isCleanTracked }
	} else {
		$invisibleCleanup += [pscustomobject]@{ Process = $process; Entry = $entry; CleanTracked = $isCleanTracked }
	}
}

Write-PreflightLog "Identified $($invisibleCleanup.Count) invisible process(es) for cleanup and $($visibleProtected.Count) visible process(es) to leave untouched."

foreach ($candidate in $invisibleCleanup) {
	$process = $candidate.Process
	try {
		$processId = [int]$process.Id
		$owner = if ($candidate.Entry) { [string]$candidate.Entry.owner } else { 'untracked' }
		$kind = if ($candidate.CleanTracked) { 'tracked clean' } else { 'rogue/untracked' }
		Write-PreflightLog "Terminating invisible $kind Excel process PID=$processId owner=$owner."
		if ($candidate.CleanTracked -and $candidate.Entry -and [string]$candidate.Entry.owner -like 'session#*') {
			$keeperPid = 0
			try { $keeperPid = [int]([string]$candidate.Entry.owner).Substring(8) } catch {}
			if ($keeperPid -gt 0) {
				try { Stop-Process -Id $keeperPid -Force -ErrorAction SilentlyContinue } catch {}
			}
		}
		Stop-Process -Id $processId -Force -ErrorAction Stop
		$exited = $false
		for ($attempt = 0; $attempt -lt 10; $attempt++) {
			Start-Sleep -Milliseconds 200
			if ($null -eq (Get-Process -Id $processId -ErrorAction SilentlyContinue)) {
				$exited = $true
				break
			}
		}

		if ($exited) {
			Write-PreflightLog "Terminated invisible $kind Excel process PID=$processId."
			if ($candidate.CleanTracked -and (Get-Command Remove-ExcelInstance -ErrorAction SilentlyContinue)) {
				Remove-ExcelInstance -InstanceId $processId -DeactivateReason 'clean-preflight'
			}
		} else {
			# A process can briefly survive Stop-Process while Windows tears down
			# the Office COM server. Retry once, still only for an invisible PID.
			$current = Get-Process -Id $processId -ErrorAction SilentlyContinue
			if ($null -ne $current -and [bool]($current.MainWindowHandle -eq 0)) {
				Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
			}
			Write-PreflightLog "Termination requested for invisible Excel process PID=$processId; process did not exit during the initial wait."
		}
	} catch {
		Write-PreflightLog "Could not terminate invisible Excel process PID=$($process.Id): $($_.Exception.Message)"
	}
}

if ($visibleProtected.Count -gt 0) {
	foreach ($candidate in $visibleProtected) {
		$process = $candidate.Process
		$kind = if ($candidate.CleanTracked) { 'tracked clean' } else { 'rogue' }
		Write-PreflightLog "Left visible $kind Excel process untouched: PID=$($process.Id), title='$($process.MainWindowTitle)'."
	}
}

$remainingInvisible = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object {
	$entry = $activeByPid[[int]$_.Id]
	$isRogue = -not $entry -or ([string]$entry.owner -match '^(rogue|unknown|unknown-owner|untracked)$')
	$isCleanTracked = $Clean -and $null -ne $entry
	($isRogue -or $isCleanTracked) -and [bool]($_.MainWindowHandle -eq 0)
})

if ($remainingInvisible.Count -gt 0) {
	Write-PreflightLog "Preflight failed: $($remainingInvisible.Count) invisible cleanup process(es) remain."
	foreach ($process in $remainingInvisible) {
		Write-PreflightLog "Invisible cleanup Excel process remains alive: PID=$($process.Id)."
	}
	exit 1
}

Write-PreflightLog 'Excel instance preflight clean.'
Write-PreflightLog 'Completed Excel instance preflight.'
exit 0
