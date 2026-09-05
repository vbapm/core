# Ensure-ExcelInstancesClean.ps1
#
# Guard run immediately before the e2e integration suite. It terminates only
# invisible rogue/untracked EXCEL.EXE processes. Visible rogue processes are
# never killed because they may belong to the user; they cause the suite to
# fail before any tests start.
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
		Stop-Process -Id $process.Id -Force -ErrorAction Stop
		Write-Output "Terminated invisible rogue Excel process PID=$($process.Id)."
	} catch {
		Write-Warning "Could not terminate invisible rogue Excel process PID=$($process.Id): $($_.Exception.Message)"
	}
}

if ($visibleRogue.Count -gt 0) {
	foreach ($process in $visibleRogue) {
		Write-Error "Visible rogue Excel process remains alive: PID=$($process.Id). It was not terminated."
	}
	exit 1
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
