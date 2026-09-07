# session.ps1
#
# Persistent PowerShell bridge for Excel COM automation.
#
# Whereas run.ps1 is a one-shot script (spawned fresh for every macro run), this
# script is a long-lived loop that reads JSON request lines from stdin, drives a
# *reused* Excel.Application instance, and prints framed JSON responses to
# stdout. Keeping the $app COM stub alive between requests lets the Node CLI
# avoid re-launching Excel for every `vba run` call, and lets later commands
# attach to the same instance.
#
# Protocol (one request per line, newline-terminated):
#   Request:  <base64-encoded JSON>
#   Response: \u001E<request-id>\u001F<json>\u001E (single line)
#
# Request JSON shape:
#   {
#     "id":       "<request-id>",       # echoed back in the response envelope
#     "appName":  "excel",
#     "file":     "<path>",
#     "macro":    "<macro name>",
#     "keepOpen": false,
#     "args":     ["<arg>", ...]
#   }
#
# Special control requests:
#   { "id": "__VBA_QUIT__" }  -> quit Excel (if we own it) and exit the loop.
#
# The instance is kept visible/hidden according to each request's background
# value. A workbook opened by a prior request that was NOT "keepOpen" is
# closed after the macro runs, but the Application instance itself is retained
# so the next request reuses it.

$ErrorActionPreference = 'Stop'

# -------
# Shared registry helpers (optional)
# -------
$RegistryModule = Join-Path $PSScriptRoot '..\scripts\ps\Excel-InstanceRegistry.ps1'
$HasRegistry = Test-Path -LiteralPath $RegistryModule
if ($HasRegistry) {
	. $RegistryModule
}

# -------
# State
# -------
$Script:App = $null
$Script:AppWasOpen = $false   # true if we attached to a pre-existing instance
$Script:BackgroundBuild = $false
$Script:OpenWorkbook = $null  # current workbook COM ref
$Script:WorkbookWasOpen = $false
$Script:IsAddin = $false      # true when the current request targets an add-in
$Script:LoadedAddins = @{}
$Script:LastRequestIsAddin = $null
$Script:WarmInstance = $false

function Get-ScriptFileName {
	param([string]$Path)
	return [System.IO.Path]::GetFileNameWithoutExtension($Path)
}

function Get-ScriptFileBase {
	param([string]$Path)
	return [System.IO.Path]::GetFileName($Path)
}

function Unescape-ScriptArgument {
	param([string]$Value)
	return $Value -replace '\^q', '"'
}

function Get-ScriptMacroTargetFile {
	param([string[]]$ArgValues)
	foreach ($arg in $ArgValues) {
		if (-not $arg) { continue }
		try {
			$parsed = $arg | ConvertFrom-Json -ErrorAction Stop
			if ($parsed.PSObject.Properties.Name -contains 'file' -and $parsed.file) {
				return ([string]$parsed.file) -replace '\\', '/'
			}
		} catch {}
	}
	return ''
}

function Get-ScriptMacroLogDescription {
	param([string]$MacroName, [string[]]$ArgValues)
	$target = Get-ScriptMacroTargetFile $ArgValues
	$log = if ($target) { "$MacroName -> $target" } else { $MacroName }
	if ($env:VBA_E2E_TEST_NAME) { $log = "$env:VBA_E2E_TEST_NAME :: $log" }
	return "$([DateTime]::UtcNow.ToString('o')) :: $log"
}

function Get-ScriptMacroErrorLog {
	param([string]$MacroName, [string[]]$ArgValues, [object]$Result, [string]$ErrorMessage = '')
	$log = Get-ScriptMacroLogDescription $MacroName $ArgValues
	if ($ErrorMessage) { return "$log :: ERROR: $ErrorMessage" }
	try {
		$parsed = $Result | ConvertFrom-Json -ErrorAction Stop
		$errors = @($parsed.errors | Where-Object { $_ })
		if ($parsed.success -eq $false -or $errors.Count -gt 0) {
			$detail = if ($errors.Count -gt 0) { $errors -join ' | ' } else { 'unknown error' }
			return "$log :: ERROR: $detail"
		}
	} catch {}
	return $log
}

function Test-ScriptWorkbookOpen {
	param([string]$WorkbookPath)
	if (-not $WorkbookPath -or $null -eq $Script:App) { return $false }
	try {
		$normalizedPath = [System.IO.Path]::GetFullPath($WorkbookPath)
		foreach ($workbook in @($Script:App.Workbooks)) {
			try {
				if ([string]::Equals([System.IO.Path]::GetFullPath([string]$workbook.FullName), $normalizedPath, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
			} catch {}
		}
	} catch {}
	return $false
}

function Refresh-ScriptInstanceLog {
	param([string]$MacroName, [string[]]$ArgValues, [object]$Result = '', [string]$ErrorMessage = '')
	if (-not $HasRegistry -or $Script:RegisteredPid -le 0 -or $null -eq $Script:App) { return }
	try {
		$log = Get-ScriptMacroErrorLog $MacroName $ArgValues $Result $ErrorMessage
		Refresh-ExcelInstance -ExcelApp $Script:App -ProcessId $Script:RegisteredPid -TargetPath (Get-ScriptMacroTargetFile $ArgValues) -Log $log
	} catch {}
}

function Write-ScriptWorkbookCloseLog {
	param([string]$WorkbookPath)

	if (-not $WorkbookPath -or -not $HasRegistry -or $Script:RegisteredPid -le 0 -or $null -eq $Script:App) {
		return
	}

	try {
		$log = "$([DateTime]::UtcNow.ToString('o')) :: Workbook.Close -> $($WorkbookPath -replace '\\', '/')"
		if ($env:VBA_E2E_TEST_NAME) {
			$log = "$([DateTime]::UtcNow.ToString('o')) :: $env:VBA_E2E_TEST_NAME :: Workbook.Close -> $($WorkbookPath -replace '\\', '/')"
		}
		Refresh-ExcelInstance -ExcelApp $Script:App -ProcessId $Script:RegisteredPid -Log $log
	} catch {
		# Registry attribution is best-effort; never mask workbook cleanup.
	}
}
$Script:RegisteredPid = 0

function Invoke-ScriptMacro {
	param(
		[object]$ExcelApp,
		[string]$MacroName,
		[string[]]$ArgValues,
		[string]$WorkbookName = '',
		[bool]$Qualified = $false
	)

	$numArgs = if ($ArgValues) { $ArgValues.Count } else { 0 }
	try {
		switch ($numArgs) {
			0  { return $ExcelApp.Run($MacroName) }
			1  { return $ExcelApp.Run($MacroName, $ArgValues[0]) }
			2  { return $ExcelApp.Run($MacroName, $ArgValues[0], $ArgValues[1]) }
			3  { return $ExcelApp.Run($MacroName, $ArgValues[0], $ArgValues[1], $ArgValues[2]) }
			4  { return $ExcelApp.Run($MacroName, $ArgValues[0], $ArgValues[1], $ArgValues[2], $ArgValues[3]) }
			5  { return $ExcelApp.Run($MacroName, $ArgValues[0], $ArgValues[1], $ArgValues[2], $ArgValues[3], $ArgValues[4]) }
			6  { return $ExcelApp.Run($MacroName, $ArgValues[0], $ArgValues[1], $ArgValues[2], $ArgValues[3], $ArgValues[4], $ArgValues[5]) }
			7  { return $ExcelApp.Run($MacroName, $ArgValues[0], $ArgValues[1], $ArgValues[2], $ArgValues[3], $ArgValues[4], $ArgValues[5], $ArgValues[6]) }
			8  { return $ExcelApp.Run($MacroName, $ArgValues[0], $ArgValues[1], $ArgValues[2], $ArgValues[3], $ArgValues[4], $ArgValues[5], $ArgValues[6], $ArgValues[7]) }
			9  { return $ExcelApp.Run($MacroName, $ArgValues[0], $ArgValues[1], $ArgValues[2], $ArgValues[3], $ArgValues[4], $ArgValues[5], $ArgValues[6], $ArgValues[7], $ArgValues[8]) }
			10 { return $ExcelApp.Run($MacroName, $ArgValues[0], $ArgValues[1], $ArgValues[2], $ArgValues[3], $ArgValues[4], $ArgValues[5], $ArgValues[6], $ArgValues[7], $ArgValues[8], $ArgValues[9]) }
		}
	} catch {
		if (-not $Qualified -and $WorkbookName) {
			$qualifiedName = "'$WorkbookName'!$MacroName"
			return Invoke-ScriptMacro $ExcelApp $qualifiedName $ArgValues $WorkbookName $true
		}
		throw
	}
	return $null
}

function Open-ScriptExcel {
	param(
		[bool]$BackgroundBuild,
		[string]$MacroName = '',
		[string[]]$ArgValues = @()
	)

	$Script:BackgroundBuild = $BackgroundBuild
	$Script:WarmInstance = $false

	if ($Script:BackgroundBuild -and $HasRegistry) {
		$reusable = Find-ReusableBackgroundExcel
		if ($null -ne $reusable) {
			$Script:App = $reusable.App
			$Script:AppWasOpen = $false
			$Script:RegisteredPid = [int]$reusable.ProcessId
			$Script:WarmInstance = $true
			try {
				Register-ExcelInstance -ExcelApp $Script:App -Owner "session#$PID" -Visible $false -Reason 'e2e' -Intent 'keep-warm' | Out-Null
			} catch {}
			Close-ScriptWorkbooks
			if ($reusable.PreviousSessionPid -gt 0 -and $reusable.PreviousSessionPid -ne $PID) {
				try { Stop-Process -Id $reusable.PreviousSessionPid -Force -ErrorAction SilentlyContinue } catch {}
			}
			return
		}
	}

	if (-not $Script:BackgroundBuild) {
		try {
			$Script:App = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
			$Script:AppWasOpen = $true
		} catch {
			# Excel not running; fall through to create a new instance
		}
	}

	if (-not $Script:AppWasOpen) {
		$lockHeld = $false
		$existingExcelPids = @{}
		foreach ($process in @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
			$existingExcelPids[[int]$process.Id] = $true
		}
		try {
			if ($HasRegistry) {
				Get-ExcelInstancesLock | Out-Null
				$lockHeld = $true
			}

			$Script:App = New-Object -ComObject "Excel.Application"
			$createdPid = 0
			if ($HasRegistry) {
				try { $createdPid = Get-ExcelProcessId -ExcelApp $Script:App } catch {}
				if ($createdPid -le 0 -or $existingExcelPids.ContainsKey($createdPid)) {
					try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Script:App) | Out-Null } catch {}
					$Script:App = $null
					throw "Refusing to automate an existing Excel process; background instance PID could not be proven new."
				}
			}
			$Script:App.Visible = if ($Script:BackgroundBuild) { $false } else { $true }
			$Script:App.ScreenUpdating = $false
			$Script:App.DisplayStatusBar = $false
			$Script:App.PrintCommunication = $false
			$Script:App.EnableAnimations = $false
			$Script:App.EnableEvents = $false
			# Match the one-shot bridge's automation behavior for every workbook
			# opened by this long-lived process. A macro-enabled workbook opened after
			# the add-in can otherwise inherit a restrictive security mode.
			try { $Script:App.AutomationSecurity = 1 } catch {}
			$Script:App.DisplayAlerts = $false

			if ($HasRegistry) {
				$reason = if ($Script:BackgroundBuild) { 'e2e' } else { 'vbapm-run' }
				try {
					$registeredPid = Register-ExcelInstance `
						-ExcelApp $Script:App `
						-Owner "session#$PID" `
						-Reason $reason `
						-Intent 'keep-warm' `
						-Log (Get-ScriptMacroLogDescription $MacroName $ArgValues)
					$Script:RegisteredPid = if ($null -ne $registeredPid) { [int]$registeredPid } else { 0 }
				} catch {
					# best-effort
				}
			}
		} finally {
			if ($lockHeld) { Release-ExcelInstancesLock }
		}
	}
}

function Find-ReusableBackgroundExcel {
	$entriesByPid = @{}
	foreach ($entry in @(Read-ExcelInstances)) {
		$entriesByPid[[int]$entry.pid] = $entry
	}

	foreach ($app in @(Get-ExcelApplications)) {
		$keepApp = $false
		try {
			$processId = Get-ExcelProcessId -ExcelApp $app
			$process = Get-Process -Id $processId -ErrorAction SilentlyContinue
			$entry = $entriesByPid[[int]$processId]
			$owner = if ($entry) { [string]$entry.owner } else { '' }
			$isTracked = $null -ne $entry -and $owner -notmatch '^(rogue|unknown|unknown-owner|untracked)$'
			$isHidden = $null -ne $process -and [bool]($process.MainWindowHandle -eq 0)
			if ($isTracked -and $isHidden) {
				$previousSessionPid = 0
				if ($owner -like 'session#*') {
					try { $previousSessionPid = [int]($owner.Substring(8)) } catch {}
				}
				$keepApp = $true
				return [pscustomobject]@{ App = $app; ProcessId = $processId; PreviousSessionPid = $previousSessionPid }
			}
		} catch {
			# Try the next ROT-reachable instance.
		}
		if (-not $keepApp) {
			try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null } catch {}
		}
	}

	return $null
}

function Ensure-ScriptWorkbook {
	param([string]$Path)

	$fileName = Get-ScriptFileName $Path
	$fullPath = [System.IO.Path]::GetFullPath($Path)

	# If the workbook from a previous request is still open and matches, reuse it.
	if ($null -ne $Script:OpenWorkbook) {
		try {
			if ($Script:OpenWorkbook.FullName -eq $fullPath) {
				return
			}
		} catch {
			$Script:OpenWorkbook = $null
		}
	}

	# Mirror run.ps1: an add-in that is already loaded must be used as-is rather
	# than reopened as a plain workbook. Reopening vbapm.xlam through
	# Workbooks.Open is what breaks unqualified Application.Run('Build.*')
	# resolution (the failure mode that caused the earlier revert).
	try {
		$addin = $Script:App.AddIns($fileName)
		if ($addin.IsOpen) {
			$Script:OpenWorkbook = $addin
			$Script:WorkbookWasOpen = $true
			return
		}
	} catch {
		# Not registered as an add-in; fall through.
	}

	# Already open *in our own instance*? Match on full path so a same-named file
	# from another directory is never picked up.
	#
	# Deliberately NOT using the registry's cross-instance Find-OpenWorkbook here:
	# re-pointing $Script:App at a foreign instance mid-session makes macro
	# resolution depend on unrelated Excel processes (and on whether that instance
	# has the add-in loaded), which is precisely how 'Cannot run the macro' shows
	# up. A session owns exactly one instance for its whole lifetime.
	try {
		foreach ($wb in $Script:App.Workbooks) {
			if ($wb.FullName -eq $fullPath) {
				$Script:OpenWorkbook = $wb
				$Script:WorkbookWasOpen = $true
				return
			}
		}
	} catch {
		# Not already open; fall through to opening it.
	}

	# Open (or reopen) the workbook in our instance.
	try {
		$Script:OpenWorkbook = $Script:App.Workbooks.Open($fullPath)
		if ($Script:BackgroundBuild) {
			$Script:App.Visible = $false
			try { $Script:App.Calculation = -4135 } catch {}
		}
	} catch {
		throw "Failed to open workbook - $($_.Exception.Message)"
	}
}

function Ensure-ScriptAddin {
	param([string]$Path)

	$fullPath = [System.IO.Path]::GetFullPath($Path)
	$fileName = Get-ScriptFileBase $Path
	$addinKey = $fullPath.ToLowerInvariant()

	if ($Script:LoadedAddins.ContainsKey($addinKey)) {
		return $Script:LoadedAddins[$addinKey]
	}

	try {
		if ($null -eq $Script:App) {
			throw 'Excel application is unavailable'
		}
		foreach ($workbook in $Script:App.Workbooks) {
			if ($workbook.FullName -eq $fullPath) {
				$Script:LoadedAddins[$addinKey] = $workbook
				return $workbook
			}
		}
		# UpdateLinks=0, ReadOnly=$true. The add-in is shared by workers, so
		# opening it read-only avoids write-lock contention on the source file.
		$workbook = $Script:App.Workbooks.Open($fullPath, 0, $true)
		if ($null -eq $workbook) {
			throw 'Excel returned no workbook object'
		}
		if ($Script:BackgroundBuild) {
			$Script:App.Visible = $false
		}
		$Script:LoadedAddins[$addinKey] = $workbook
		return $workbook
	} catch {
		$fallbackError = $_.Exception.Message
		throw "Failed to load add-in '$fileName' - $fallbackError"
	}
}

function Close-ScriptWorkbooks {
	try {
		foreach ($workbook in @($Script:App.Workbooks)) {
			$displayPath = ''
			try { $displayPath = [string]$workbook.FullName } catch {}
			$fullPath = ''
			try { $fullPath = [System.IO.Path]::GetFullPath($displayPath).ToLowerInvariant() } catch {}

			$isLoadedAddin = $false
			if ($fullPath) {
				$isLoadedAddin = $Script:LoadedAddins.ContainsKey($fullPath)
			}
			if ($isLoadedAddin -or $fullPath -match '\.(xlam|xla)$') {
				continue
			}

			$closed = $false
			try { $workbook.Close($true); $closed = $true } catch {}
			if ($closed) { Write-ScriptWorkbookCloseLog $displayPath }
			try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
		}
	} catch {
		# Best-effort cleanup; the next request can still report its own error.
	}

	# Add-ins are needed only for the request that invoked them. Keeping one
	# loaded changes the macro context for a later workbook request, so unload
	# session-owned add-ins while retaining the Excel.Application itself.
	foreach ($addin in @($Script:LoadedAddins.Values)) {
		$displayPath = ''
		try { $displayPath = [string]$addin.FullName } catch {}
		try { $addin.Installed = $false } catch {}
		$closed = $false
		try { $addin.Close($false); $closed = $true } catch {}
		if ($closed) { Write-ScriptWorkbookCloseLog $displayPath }
		try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($addin) | Out-Null } catch {}
	}
	$Script:LoadedAddins = @{}
	$Script:OpenWorkbook = $null
	$Script:WorkbookWasOpen = $false
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers()
}

function Reset-ScriptExcel {
	$lockHeld = $false
	try {
		if ($HasRegistry) {
			Get-ExcelInstancesLock | Out-Null
			$lockHeld = $true
		}
		Close-ScriptWorkbooks
		if ($null -ne $Script:App) {
			try { $Script:App.Quit() } catch {}
			try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Script:App) | Out-Null } catch {}
		}
	} finally {
		if ($lockHeld) { Release-ExcelInstancesLock }
	}
	$Script:App = $null
	$Script:AppWasOpen = $false
	$Script:LoadedAddins = @{}
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers()
}

function Reset-ScriptWorkbooks {
	if ($null -eq $Script:App) {
		return
	}

	Close-ScriptWorkbooks
	$Script:IsAddin = $false
	$Script:LastRequestIsAddin = $null
}

function Ensure-WarmAnchor {
	if ($null -eq $Script:App) { return }

	$anchorPath = Join-Path $env:TEMP 'Excel-Instances\warm-anchor.xlsx'
	try {
		if (Test-Path -LiteralPath $anchorPath) {
			Remove-Item -LiteralPath $anchorPath -Force -ErrorAction SilentlyContinue
		}
		$anchor = $Script:App.Workbooks.Add()
		$anchor.SaveAs($anchorPath, 51)
		$anchor.Close($true)
		$Script:App.Workbooks.Open($anchorPath, 0, $true) | Out-Null
	} catch {
		# A warm anchor is an optimization; the Excel instance remains usable
		# without it when Excel cannot save the temporary workbook.
	}
}

function Invoke-ScriptRun {
	param(
		[string]$FilePath,
		[string]$MacroName,
		[bool]$KeepOpen,
		[bool]$Background,
		[string[]]$ArgValues
	)

	$requestIsAddin = $FilePath -match '\.(xlam|xla)$'
	if ($null -ne $Script:App -and $Script:BackgroundBuild -ne $Background) {
		Reset-ScriptExcel
	}
	if ($null -ne $Script:App) {
		try {
			$null = $Script:App.Workbooks.Count
		} catch {
			Reset-ScriptExcel
		}
	}
	if ($null -ne $Script:App -and $null -ne $Script:LastRequestIsAddin -and $Script:LastRequestIsAddin -ne $requestIsAddin) {
		# Excel does not reliably return to a clean VBA project context when a
		# request changes from an add-in to a macro-enabled workbook (or back).
		# Reuse remains enabled for consecutive requests of the same kind, while
		# context switches get a fresh owned Excel instance.
		Reset-ScriptExcel
	}

	if ($null -eq $Script:App) {
		Open-ScriptExcel $Background $MacroName $ArgValues
	}
	$targetPath = Get-ScriptMacroTargetFile $ArgValues
	$targetWasOpen = $MacroName -eq 'Build.CloseFile' -and (Test-ScriptWorkbookOpen $targetPath)

	# Per-request state. $Script:WorkbookWasOpen describes *this* request only;
	# leaving it sticky from a previous request means we stop closing workbooks we
	# opened, so they pile up in the reused instance and bleed into later macros.
	$Script:WorkbookWasOpen = $false
	$Script:IsAddin = $requestIsAddin
	$Script:LastRequestIsAddin = $requestIsAddin

	if ($Script:IsAddin) {
		$Script:OpenWorkbook = Ensure-ScriptAddin $FilePath
		$Script:WorkbookWasOpen = $true
	} else {
		Ensure-ScriptWorkbook $FilePath
	}

	# A previous request may have left the add-in or another workbook active.
	# Application.Run resolves unqualified macro names against the active VBA
	# project, so make the current request's workbook the active document first.
	try { $Script:OpenWorkbook.Activate() } catch {}

	$workbookName = if ($Script:IsAddin) { '' } else { Get-ScriptFileBase $FilePath }
	try {
		$result = Invoke-ScriptMacro $Script:App $MacroName $ArgValues $workbookName
		if ($targetWasOpen -and -not (Test-ScriptWorkbookOpen $targetPath)) {
			Write-ScriptWorkbookCloseLog $targetPath
		}
	} catch {
		# Excel can retain an unusable VBA project context after a macro opened
		# and closed workbooks internally. Recreate only the owned application and
		# retry once; healthy request sequences still reuse the same process.
		Reset-ScriptExcel
		Open-ScriptExcel $Background $MacroName $ArgValues
		$Script:WorkbookWasOpen = $false
		if ($Script:IsAddin) {
			$Script:OpenWorkbook = Ensure-ScriptAddin $FilePath
			$Script:WorkbookWasOpen = $true
		} else {
			Ensure-ScriptWorkbook $FilePath
		}
		try { $Script:OpenWorkbook.Activate() } catch {}
		$result = Invoke-ScriptMacro $Script:App $MacroName $ArgValues $workbookName
	}

	# VBA build macros open their target workbook internally. Keep those internal
	# workbooks alive until the add-in-to-workbook context reset, matching the
	# one-shot bridge's lifecycle; closing them immediately can discard the VBA
	# project before the target archive is finalized.
	if (-not $KeepOpen -and -not $Script:IsAddin) {
		Close-ScriptWorkbooks
	}

	return $result
}

function Wait-ScriptExcelExit {
	param(
		[int]$ProcessId,
		[int]$TimeoutSeconds = 10
	)

	if ($ProcessId -le 0) {
		return $true
	}

	$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
	while ((Get-Date) -lt $deadline) {
		if (-not (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)) {
			return $true
		}
		Start-Sleep -Milliseconds 100
	}

	return (-not (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue))
}

function Test-KeepWarmExcel {
	if (-not $Script:BackgroundBuild -or $env:E2E_BACKGROUND -notmatch '^(1|true|yes)$') {
		return $false
	}

	try {
		$invisible = @(Get-RunningExcelInstances | Where-Object { -not $_.visible })
		return $invisible.Count -le 1
	} catch {
		return $false
	}
}

function Close-ScriptSession {
	$lockHeld = $false
	try {
		if ($HasRegistry) {
			Get-ExcelInstancesLock | Out-Null
			$lockHeld = $true
		}
		if ($null -ne $Script:App) {
			Close-ScriptWorkbooks
		}
		$keepWarm = Test-KeepWarmExcel
		# Quit Excel only if we launched it ourselves.
		if ($keepWarm -and $null -ne $Script:App) {
			Write-InstanceLog "detaching warm Excel instance excelPid=$($Script:RegisteredPid)"
			try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Script:App) | Out-Null } catch {}
		} elseif (-not $Script:AppWasOpen -and $null -ne $Script:App) {
			$excelPid = 0
			if (Get-Command Get-ExcelProcessId -ErrorAction SilentlyContinue) {
				try { $excelPid = Get-ExcelProcessId -ExcelApp $Script:App } catch {}
			}
			try {
				$Script:App.Quit()
			} catch {
				# best-effort
			}
			try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Script:App) | Out-Null } catch {}
			[System.GC]::Collect()
			[System.GC]::WaitForPendingFinalizers()

			# A COM Quit can return while Excel is still alive, or throw after Excel
			# has become unresponsive. This PID was captured from the instance we
			# created, so force-cleaning it cannot touch a user's visible Excel.
			if ($excelPid -gt 0) {
				try {
					if (Get-Process -Id $excelPid -ErrorAction SilentlyContinue) {
						Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
					}
				} catch {}

				# Keep the registry lock while the process exits and COM/file handles
				# settle. A new worker must not create Excel in this interval.
				$exited = Wait-ScriptExcelExit -ProcessId $excelPid
				if ($exited) {
					Start-Sleep -Seconds 2
				}

				if (Get-Command Unregister-ExcelInstance -ErrorAction SilentlyContinue) {
					try { Unregister-ExcelInstance -ProcessId $excelPid } catch {}
				}
			}
		}
	} finally {
		if ($lockHeld) { Release-ExcelInstancesLock }
	}
	$Script:App = $null
	$Script:RegisteredPid = 0
	$Script:WarmInstance = $false
}

# -------
# Framed request loop
# -------

function Write-ScriptResponse {
	param([string]$RequestId, [string]$Json, [object]$Writer = $null)

	$line = [char]0x1E + $RequestId + [char]0x1F + $Json
	if ($null -eq $Writer) {
		[Console]::Out.WriteLine($line)
		[Console]::Out.Flush()
	} else {
		$Writer.WriteLine($line)
		$Writer.Flush()
	}
}

function Invoke-ScriptRequest {
	param([object]$RequestJson, [scriptblock]$Respond)

	$requestId = if ($RequestJson.id) { [string]$RequestJson.id } else { 'unknown' }
	if ($requestId -eq '__VBA_QUIT__') {
		Close-ScriptSession
		& $Respond $requestId (@{ success = $true; result = 'quit' } | ConvertTo-Json -Compress)
		return 'quit'
	}
	if ($requestId -eq '__VBA_RESET__') {
		try { Reset-ScriptWorkbooks; & $Respond $requestId (@{ success = $true; result = 'reset' } | ConvertTo-Json -Compress) }
		catch { & $Respond $requestId (@{ success = $false; errors = @($_.Exception.Message) } | ConvertTo-Json -Compress) }
		return 'continue'
	}
	if ($requestId -eq '__VBA_DETACH__') {
		try {
			Reset-ScriptWorkbooks
			& $Respond $requestId (@{ success = $true; result = 'detached' } | ConvertTo-Json -Compress)
		} catch { & $Respond $requestId (@{ success = $false; errors = @($_.Exception.Message) } | ConvertTo-Json -Compress) }
		return 'detach'
	}

	$resolveAppName = if ($RequestJson.appName) { [string]$RequestJson.appName } else { 'excel' }
	if ($resolveAppName -ne 'excel') {
		& $Respond $requestId (@{ success = $false; errors = @('Unsupported app') } | ConvertTo-Json -Compress)
		return 'continue'
	}
	$resolveFile = [string]$RequestJson.file
	$resolveMacro = [string]$RequestJson.macro
	$previousTestName = $env:VBA_E2E_TEST_NAME
	if ($RequestJson.testName) { $env:VBA_E2E_TEST_NAME = [string]$RequestJson.testName }
	$resolveKeepOpen = [bool]$RequestJson.keepOpen
	$resolveBackground = [bool]$RequestJson.background
	$resolveArgs = @()
	if ($RequestJson.args) { foreach ($a in $RequestJson.args) { $resolveArgs += Unescape-ScriptArgument ([string]$a) } }
	try {
		$result = Invoke-ScriptRun $resolveFile $resolveMacro $resolveKeepOpen $resolveBackground $resolveArgs
		Refresh-ScriptInstanceLog -MacroName $resolveMacro -ArgValues $resolveArgs -Result $result
		& $Respond $requestId (@{ success = $true; result = $result } | ConvertTo-Json -Compress)
	} catch {
		Refresh-ScriptInstanceLog -MacroName $resolveMacro -ArgValues $resolveArgs -ErrorMessage $_.Exception.Message
		& $Respond $requestId (@{ success = $false; errors = @($_.Exception.Message) } | ConvertTo-Json -Compress)
	} finally {
		if ($null -eq $previousTestName) { Remove-Item Env:VBA_E2E_TEST_NAME -ErrorAction SilentlyContinue }
		else { $env:VBA_E2E_TEST_NAME = $previousTestName }
	}
	return 'continue'
}

function Invoke-ScriptRequestLine {
	param([string]$Line, [scriptblock]$Respond)
	try {
		$decoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Line.Trim()))
		$requestJson = $decoded | ConvertFrom-Json
	} catch {
		& $Respond 'error' (@{ success = $false; errors = @("Invalid request: $($_.Exception.Message)") } | ConvertTo-Json -Compress)
		return 'continue'
	}
	return Invoke-ScriptRequest -RequestJson $requestJson -Respond $Respond
}

function Start-WarmSessionServer {
	$metadataPath = Join-Path $env:TEMP 'Excel-Instances\warm-session.json'
	$listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
	$listener.Start()
	$port = ([System.Net.IPEndPoint]$listener.LocalEndpoint).Port
	[System.IO.File]::WriteAllText($metadataPath, (@{ port = $port; pid = $PID } | ConvertTo-Json -Compress), [System.Text.UTF8Encoding]::new($false))
	try {
		while ($true) {
			$client = $listener.AcceptTcpClient()
			$reader = New-Object System.IO.StreamReader($client.GetStream())
			$writer = New-Object System.IO.StreamWriter($client.GetStream())
			$writer.AutoFlush = $true
			$state = 'continue'
			try {
				while ($null -ne ($line = $reader.ReadLine())) {
					$state = Invoke-ScriptRequestLine -Line $line -Respond { param($id, $json) Write-ScriptResponse $id $json $writer }
					if ($state -eq 'quit' -or $state -eq 'detach') { break }
				}
			} finally { $reader.Dispose(); $writer.Dispose(); $client.Close() }
			if ($state -eq 'quit') { break }
		}
	} finally {
		$listener.Stop()
		Remove-Item -LiteralPath $metadataPath -Force -ErrorAction SilentlyContinue
	}
}

while ($true) {
	$line = [Console]::In.ReadLine()
	if ($null -eq $line) { Write-InstanceLog 'persistent session stdin reached EOF'; break }
	if ($line.Trim().Length -eq 0) { continue }
	$state = Invoke-ScriptRequestLine -Line $line -Respond { param($id, $json) Write-ScriptResponse $id $json }
	if ($state -eq 'quit') { break }
	if ($state -eq 'detach') { Start-WarmSessionServer; break }
}

if ($state -ne 'quit') { Close-ScriptSession }
