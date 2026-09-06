# Ensure-ExcelInstancesClean.ps1
#
# Guard run immediately before the e2e integration suite. It terminates only
# invisible rogue/untracked EXCEL.EXE processes. Visible rogue processes are
# never killed because they may belong to the user; they are reported as
# warnings while the suite continues.
#
# The e2e suite itself creates *background* instances (VBA_BACKGROUND_BUILD=1)
# and registers each one, so in the common case there is nothing running when
# this guard fires.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\scripts\ps\Ensure-ExcelInstancesClean.ps1

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'Excel-InstanceRegistry.ps1')

Sync-ExcelInstanceSnapshot | Out-Null
$activeByPid = @{}
foreach ($entry in @(Read-ExcelInstances)) {
	$activeByPid[[int]$entry.pid] = $entry
}

$visibleRogue = @()
$invisibleRogue = @()
foreach ($process in @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
	$entry = $activeByPid[[int]$process.Id]
	$owner = if ($entry) { [string]$entry.owner } else { 'untracked' }
	$isRogue = -not $entry -or $owner -match '^(rogue|unknown|unknown-owner|untracked)$'
	if (-not $isRogue) { continue }

	if ([bool]($process.MainWindowHandle -ne 0)) {
		$visibleRogue += $process
	} else {
		$invisibleRogue += $process
	}
}

foreach ($process in $invisibleRogue) {
	try {
		$processId = [int]$process.Id
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
			Write-Output "Terminated invisible rogue Excel process PID=$processId."
		} else {
			# A process can briefly survive Stop-Process while Windows tears down
			# the Office COM server. Retry once, still only for an invisible PID.
			$current = Get-Process -Id $processId -ErrorAction SilentlyContinue
			if ($null -ne $current -and [bool]($current.MainWindowHandle -eq 0)) {
				Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
			}
			Write-Output "Termination requested for invisible rogue Excel process PID=$processId."
		}
	} catch {
		Write-Warning "Could not terminate invisible rogue Excel process PID=$($process.Id): $($_.Exception.Message)"
	}
}

if ($visibleRogue.Count -gt 0) {
	foreach ($process in $visibleRogue) {
		Write-Warning "Visible rogue Excel process left untouched: PID=$($process.Id), title='$($process.MainWindowTitle)'."
	}
}

$remainingInvisible = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object {
	$entry = $activeByPid[[int]$_.Id]
	$isRogue = -not $entry -or ([string]$entry.owner -match '^(rogue|unknown|unknown-owner|untracked)$')
	$isRogue -and [bool]($_.MainWindowHandle -eq 0)
})

if ($remainingInvisible.Count -gt 0) {
	foreach ($process in $remainingInvisible) {
		Write-Error "Invisible rogue Excel process remains alive: PID=$($process.Id)."
	}
	exit 1
}

Write-Output 'Excel instance preflight clean.'
exit 0
