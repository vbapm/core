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
	param([bool]$BackgroundBuild)

	$Script:BackgroundBuild = $BackgroundBuild

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
		try {
			if ($HasRegistry) {
				Get-ExcelInstancesLock | Out-Null
				$lockHeld = $true
			}

			$Script:App = New-Object -ComObject "Excel.Application"
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
					Register-ExcelInstance -ExcelApp $Script:App -Owner "session#$PID" -Reason $reason | Out-Null
				} catch {
					# best-effort
				}
			}
		} finally {
			if ($lockHeld) { Release-ExcelInstancesLock }
		}
	}
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
			$fullPath = ''
			try { $fullPath = [System.IO.Path]::GetFullPath([string]$workbook.FullName).ToLowerInvariant() } catch {}

			$isLoadedAddin = $false
			if ($fullPath) {
				$isLoadedAddin = $Script:LoadedAddins.ContainsKey($fullPath)
			}
			if ($isLoadedAddin -or $fullPath -match '\.(xlam|xla)$') {
				continue
			}

			try { $workbook.Close($true) } catch {}
			try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
		}
	} catch {
		# Best-effort cleanup; the next request can still report its own error.
	}

	# Add-ins are needed only for the request that invoked them. Keeping one
	# loaded changes the macro context for a later workbook request, so unload
	# session-owned add-ins while retaining the Excel.Application itself.
	foreach ($addin in @($Script:LoadedAddins.Values)) {
		try { $addin.Installed = $false } catch {}
		try { $addin.Close($false) } catch {}
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
		Open-ScriptExcel $Background
	}

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
	} catch {
		# Excel can retain an unusable VBA project context after a macro opened
		# and closed workbooks internally. Recreate only the owned application and
		# retry once; healthy request sequences still reuse the same process.
		Reset-ScriptExcel
		Open-ScriptExcel $Background
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
		# Quit Excel only if we launched it ourselves.
		if (-not $Script:AppWasOpen -and $null -ne $Script:App) {
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
}

# -------
# Framed request loop
# -------

function Write-ScriptResponse {
	param([string]$RequestId, [string]$Json)

	# Envelope: \u001E<id>\u001F<json>  (single line, record separator delimiters)
	[Console]::Out.WriteLine([char]0x1E + $RequestId + [char]0x1F + $Json)
	[Console]::Out.Flush()
}

while ($true) {
	$line = [Console]::In.ReadLine()
	if ($null -eq $line) {
		break
	}
	if ($line.Trim().Length -eq 0) {
		continue
	}

	$requestJson = $null
	try {
		$decoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($line.Trim()))
		$requestJson = $decoded | ConvertFrom-Json
	} catch {
		Write-ScriptResponse 'error' (@{ success = $false; errors = @("Invalid request: $($_.Exception.Message)") } | ConvertTo-Json -Compress)
		continue
	}

	$requestId = if ($requestJson.id) { [string]$requestJson.id } else { 'unknown' }

	# Control: quit.
	if ($requestId -eq '__VBA_QUIT__') {
		Close-ScriptSession
		break
	}

	$resolveAppName = if ($requestJson.appName) { [string]$requestJson.appName } else { 'excel' }
	if ($resolveAppName -ne 'excel') {
		Write-ScriptResponse $requestId (@{ success = $false; errors = @('Unsupported app') } | ConvertTo-Json -Compress)
		continue
	}

	$resolveFile = [string]$requestJson.file
	$resolveMacro = [string]$requestJson.macro
	$resolveKeepOpen = [bool]$requestJson.keepOpen
	$resolveBackground = [bool]$requestJson.background
	$resolveArgs = @()
	if ($requestJson.args) {
		foreach ($a in $requestJson.args) {
			$resolveArgs += Unescape-ScriptArgument ([string]$a)
		}
	}

	try {
		$result = Invoke-ScriptRun $resolveFile $resolveMacro $resolveKeepOpen $resolveBackground $resolveArgs
		$payload = @{ success = $true; result = $result } | ConvertTo-Json -Compress
		Write-ScriptResponse $requestId $payload
	} catch {
		$payload = @{ success = $false; errors = @($_.Exception.Message) } | ConvertTo-Json -Compress
		Write-ScriptResponse $requestId $payload
	}
}

# Reaching here means either an explicit __VBA_QUIT__ (already cleaned up, and
# Close-ScriptSession is a no-op the second time) or stdin hit EOF because the
# parent Node process exited without sending one. Always quit the instance we
# own, otherwise a killed/crashed caller strands a hidden EXCEL.EXE forever.
Close-ScriptSession
