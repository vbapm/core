# run.ps1
#
# One-shot PowerShell bridge for Excel COM automation in FOREGROUND mode.
#
# This is the Windows counterpart of run.applescript (see the feature-parity
# notes in that file). It attaches to an already-running (visible) Excel
# instance when one exists, otherwise creates a new visible one, runs the macro,
# and closes/quits only what it opened. Background (hidden-instance) automation
# lives in run-background.ps1 so this file stays a tight mirror of its macOS
# counterpart.
#
# Invoked as (AppName, File, Command, ...MacroArgs) plus an optional -KeepOpen
# switch. There is no -Background switch — this script is foreground-only.

param(
	[Parameter(Position=0)]
	[string]$AppName,

	[Parameter(Position=1)]
	[string]$File,

	[Parameter(Position=2)]
	[string]$Command,

	[switch]$KeepOpen,

	[Parameter(Position=3, ValueFromRemainingArguments=$true)]
	[string[]]$MacroArgs
)

# -------
# Helpers
# -------

function Unescape {
	param([string]$Value)
	return $Value -replace '\^q', '"'
}

function GetFileBase {
	param([string]$Path)
	return [System.IO.Path]::GetFileName($Path)
}

function GetFileName {
	param([string]$Path)
	return [System.IO.Path]::GetFileNameWithoutExtension($Path)
}

# Extract the target `file` path from a macro's JSON argument (if present).
# build/export/extract open that file *inside* the VBA macro (run.ps1 itself only
# opens the vbapm.xlam add-in), so this lets the registry attribute an instance
# to the file a command actually operated on.
function Get-MacroTargetFile {
	param([string[]]$MacroArgValues)

	foreach ($arg in $MacroArgValues) {
		if (-not $arg) { continue }
		try {
			$parsed = $arg | ConvertFrom-Json -ErrorAction Stop
			if ($null -ne $parsed -and $parsed.PSObject.Properties.Name -contains 'file' -and $parsed.file) {
				return ([string]$parsed.file) -replace '\\', '/'
			}
		} catch {
			# Not JSON, or no `file` field — skip.
		}
	}

	return ''
}

function Get-MacroLogDescription {
	param(
		[string]$MacroName,
		[string[]]$MacroArgValues
	)

	$testName = $env:VBA_E2E_TEST_NAME
	$target = Get-MacroTargetFile $MacroArgValues
	$log = if ($target) { "$MacroName -> $target" } else { $MacroName }
	if ($testName) { $log = "$testName :: $log" }
	return "$([DateTime]::UtcNow.ToString('o')) :: $log"
}

function Get-MacroErrorLog {
	param(
		[string]$MacroName,
		[string[]]$MacroArgValues,
		[string]$Result
	)

	$log = Get-MacroLogDescription $MacroName $MacroArgValues
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

function Write-WorkbookCloseLog {
	param(
		[object]$ExcelApp,
		[int]$ProcessId,
		[string]$WorkbookPath
	)

	if (-not $WorkbookPath -or -not (Get-Command Refresh-ExcelInstance -ErrorAction SilentlyContinue)) {
		return
	}

	try {
		$testName = $env:VBA_E2E_TEST_NAME
		$log = "Workbook.Close -> $($WorkbookPath -replace '\\', '/')"
		if ($testName) { $log = "$testName :: $log" }
		$log = "$([DateTime]::UtcNow.ToString('o')) :: $log"
		Refresh-ExcelInstance -ExcelApp $ExcelApp -ProcessId $ProcessId -Log $log
	} catch {
		# Registry attribution is best-effort; never mask workbook cleanup.
	}
}

function Test-WorkbookOpen {
	param(
		[object]$ExcelApp,
		[string]$WorkbookPath
	)

	if ($null -eq $ExcelApp -or -not $WorkbookPath) { return $false }
	try {
		$normalizedPath = [System.IO.Path]::GetFullPath($WorkbookPath)
		foreach ($workbook in @($ExcelApp.Workbooks)) {
			try {
				if ([string]::Equals([System.IO.Path]::GetFullPath([string]$workbook.FullName), $normalizedPath, [System.StringComparison]::OrdinalIgnoreCase)) {
					return $true
				}
			} catch {}
		}
	} catch {}
	return $false
}

function Fail {
	param([string]$Message)

	PrintLn "{`"success`":false,`"errors`":[`"$Message`"]}"
	exit 1
}

function Print {
	param([string]$Message)
	[Console]::Out.Write($Message)
}

function PrintLn {
	param([string]$Message)
	[Console]::Out.WriteLine($Message)
}

function PrintErr {
	param([string]$Message)
	[Console]::Error.Write($Message)
}

function Write-BridgeLog {
	param([string]$Message)

	if (Get-Command Write-InstanceLog -ErrorAction SilentlyContinue) {
		Write-InstanceLog "bridge=run.ps1 $Message"
	}
}

# Instance-lifecycle debug logging (`Get-InstanceLogPath` / `Write-InstanceLog`,
# gated on VBA_DEBUG_INSTANCES) lives in scripts/ps/Excel-InstanceRegistry.ps1
# alongside the rest of the instance-coordination machinery it reports on.
#
# A "run banner" (clear separator marking the start of a new e2e run) is written
# by the test-side globalSetup before tests start, since `run.ps1` is a
# short-lived process per macro call and can't detect run boundaries.

# -------
# Run Macro
# -------

# Note: Do NOT name parameters "$Args" — it conflicts with the PowerShell
# automatic variable $args and can silently become an empty array.

function RunMacro {
	param(
		[object]$ExcelApp,
		[string]$MacroName,
		[string[]]$MacroArgValues
	)

	$numArgs = if ($MacroArgValues) { $MacroArgValues.Count } else { 0 }
	switch ($numArgs) {
		0  { return $ExcelApp.Run($MacroName) }
		1  { return $ExcelApp.Run($MacroName, $MacroArgValues[0]) }
		2  { return $ExcelApp.Run($MacroName, $MacroArgValues[0], $MacroArgValues[1]) }
		3  { return $ExcelApp.Run($MacroName, $MacroArgValues[0], $MacroArgValues[1], $MacroArgValues[2]) }
		4  { return $ExcelApp.Run($MacroName, $MacroArgValues[0], $MacroArgValues[1], $MacroArgValues[2], $MacroArgValues[3]) }
		5  { return $ExcelApp.Run($MacroName, $MacroArgValues[0], $MacroArgValues[1], $MacroArgValues[2], $MacroArgValues[3], $MacroArgValues[4]) }
		6  { return $ExcelApp.Run($MacroName, $MacroArgValues[0], $MacroArgValues[1], $MacroArgValues[2], $MacroArgValues[3], $MacroArgValues[4], $MacroArgValues[5]) }
		7  { return $ExcelApp.Run($MacroName, $MacroArgValues[0], $MacroArgValues[1], $MacroArgValues[2], $MacroArgValues[3], $MacroArgValues[4], $MacroArgValues[5], $MacroArgValues[6]) }
		8  { return $ExcelApp.Run($MacroName, $MacroArgValues[0], $MacroArgValues[1], $MacroArgValues[2], $MacroArgValues[3], $MacroArgValues[4], $MacroArgValues[5], $MacroArgValues[6], $MacroArgValues[7]) }
		9  { return $ExcelApp.Run($MacroName, $MacroArgValues[0], $MacroArgValues[1], $MacroArgValues[2], $MacroArgValues[3], $MacroArgValues[4], $MacroArgValues[5], $MacroArgValues[6], $MacroArgValues[7], $MacroArgValues[8]) }
		10 { return $ExcelApp.Run($MacroName, $MacroArgValues[0], $MacroArgValues[1], $MacroArgValues[2], $MacroArgValues[3], $MacroArgValues[4], $MacroArgValues[5], $MacroArgValues[6], $MacroArgValues[7], $MacroArgValues[8], $MacroArgValues[9]) }
	}

	return $null
}

# -------
# Excel instance coordination registry
# -------

# Track our created Excel instance in %TEMP%\Excel-Instances so concurrent
# agents can coordinate. Dot-source the shared registry module; if it isn't
# present (e.g. installed CLI without scripts/), coordination is a no-op.
$RegistryModule = Join-Path $PSScriptRoot '..\scripts\ps\Excel-InstanceRegistry.ps1'
$HasRegistry = Test-Path -LiteralPath $RegistryModule
if ($HasRegistry) {
	. $RegistryModule
}

# -------
# Excel
# -------

class Excel {
	hidden [object]$App
	hidden [bool]$ExcelWasOpen = $false
	hidden [object]$Workbook
	hidden [bool]$WorkbookWasOpen = $false
	hidden [bool]$IsAddin = $false
	hidden [int]$Pid = 0

	Excel() {
		$this.OpenExcel()
	}

	Excel([object]$App, [object]$Workbook) {
		$this.App = $App
		$this.ExcelWasOpen = $true
		$this.Workbook = $Workbook
		$this.WorkbookWasOpen = $true
		Write-BridgeLog "attach=target-workbook appWasOpen=true"
	}

	Excel([object]$App) {
		$this.App = $App
		$this.ExcelWasOpen = $true
		Write-BridgeLog "attach=target-application appWasOpen=true"
	}

	[string] Run([string]$FilePath, [string]$MacroName, [string[]]$MacroArgValues) {
		$this.OpenWorkbook($FilePath)
		$targetPath = Get-MacroTargetFile $MacroArgValues
		$targetWasOpen = $MacroName -eq 'Build.CloseFile' -and (Test-WorkbookOpen -ExcelApp $this.App -WorkbookPath $targetPath)

		# Record the just-opened workbook/addin in the coordination registry NOW,
		# while the COM object is still in memory and the workbook is open, and
		# also record the file the macro will operate on (build/export/extract
		# open it inside the macro). Capturing at open time (rather than only at
		# teardown) keeps the workbook list accurate even when the macro closes
		# the workbook before the run returns — which is what lets us attribute a
		# lingering instance to the exact test that opened the workbook.
		if (Get-Command Refresh-ExcelInstance -ErrorAction SilentlyContinue) {
			try {
				$log = if ($this.ExcelWasOpen) { '' } else { Get-MacroLogDescription $MacroName $MacroArgValues }
				Refresh-ExcelInstance -ExcelApp $this.App -ProcessId $this.Pid -TargetPath (Get-MacroTargetFile $MacroArgValues) -Log $log
			} catch {
				# best-effort; never break a run over coordination bookkeeping
			}
		}

		$result = RunMacro $this.App $MacroName $MacroArgValues
		if ($targetWasOpen -and -not (Test-WorkbookOpen -ExcelApp $this.App -WorkbookPath $targetPath)) {
			Write-WorkbookCloseLog -ExcelApp $this.App -ProcessId $this.Pid -WorkbookPath $targetPath
		}

		return $result
	}

	hidden [void] OpenExcel() {
		# Foreground mode: attach to an already-running (visible) Excel if there
		# is one, otherwise create a new visible instance.
		Write-BridgeLog "open-request background=false"

		try {
			$this.App = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
			$this.ExcelWasOpen = $true
			Write-BridgeLog "attach=get-active-object appWasOpen=true"
		} catch {
			Write-BridgeLog "attach=get-active-object failed error=$($_.Exception.Message)"
			# Excel not running; fall through to create a new instance
		}

		if (-not $this.ExcelWasOpen) {
			Write-BridgeLog "create-start background=false"
			$existingExcelPids = @{}
			foreach ($process in @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
				$existingExcelPids[[int]$process.Id] = $true
			}

			try {
				$this.App = New-Object -ComObject "Excel.Application"
				if (Get-Command Get-ExcelProcessId -ErrorAction SilentlyContinue) {
					$createdPid = 0
					try { $createdPid = Get-ExcelProcessId -ExcelApp $this.App } catch {}
					if ($createdPid -le 0 -or $existingExcelPids.ContainsKey($createdPid)) {
						try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.App) | Out-Null } catch {}
						$this.App = $null
						throw "Refusing to automate an existing Excel process; new instance PID could not be proven."
					}
				}
				$this.App.Visible = $true
				$this.App.ScreenUpdating = $false
				$this.App.DisplayStatusBar = $false
				$this.App.PrintCommunication = $false
				$this.App.EnableAnimations = $false
				$this.App.EnableEvents = $false
				# Suppress prompts (e.g. "save changes?" on a peer addin loaded
				# via References.AddFromFile) so Close()/Quit() don't block and
				# leave a lingering Excel process holding file locks.
				$this.App.DisplayAlerts = $false
				Write-BridgeLog "create-complete visible=$($this.App.Visible)"
			} catch {
				Fail "ERROR #5: Failed to open Excel - $($_.Exception.Message)"
			}

			$excelPid = 0
			try { $excelPid = Get-ExcelProcessId -ExcelApp $this.App } catch { $excelPid = 0 }
			$this.Pid = $excelPid
			Write-InstanceLog "created Excel instance excelPid=$excelPid mode=foreground"
		}
	}

	hidden [void] OpenWorkbook([string]$Path) {
		$fileName = GetFileName $Path
		$fileBase = GetFileBase $Path
		$fullPath = [System.IO.Path]::GetFullPath($Path)

		# Only the vbapm add-in is left open across runs (avoids close/reopen
		# churn). Any other add-in or workbook closes normally after the run.
		$this.IsAddin = $fileBase -eq 'vbapm.xlam'

		# If we already attached to a workbook (found open in another instance),
		# skip the open logic entirely.
		if ($this.WorkbookWasOpen -and $null -ne $this.Workbook) {
			return
		}

		# Check add-ins first
		try {
			$addin = $this.App.AddIns($fileName)
			if ($addin.IsOpen) {
				$this.Workbook = $addin
				$this.WorkbookWasOpen = $true
				return
			}
		} catch {
			# Not found in add-ins, continue
		}

		# Check already-open workbooks — match by full path to avoid picking up
		# a same-named file from a different directory.
		try {
			foreach ($wb in $this.App.Workbooks) {
				if ($wb.FullName -eq $fullPath) {
					$this.Workbook = $wb
					$this.WorkbookWasOpen = $true
					return
				}
			}
		} catch {
			# Not already open, continue
		}

		# Open the workbook
		try {
			$this.Workbook = $this.App.Workbooks.Open($fullPath)
		} catch {
			Fail "ERROR #6: Failed to open workbook - $($_.Exception.Message)"
		}
	}

	[void] Dispose([bool]$KeepOpen) {
		Write-BridgeLog "dispose-enter appWasOpen=$($this.ExcelWasOpen) isAddin=$($this.IsAddin) workbookWasOpen=$($this.WorkbookWasOpen) keepOpen=$KeepOpen"

		# An add-in we opened is left open for reuse across runs (we're reusing
		# the user's already-running Excel).
		if ($this.IsAddin) {
			Write-BridgeLog "dispose-skip addin-foreground"
			return
		}

		# A file that was open before we started is never closed by us
		$closeWorkbook = -not $this.WorkbookWasOpen -and -not $KeepOpen

		if ($closeWorkbook -and $null -ne $this.Workbook) {
			$workbookPath = ''
			try { $workbookPath = [string]$this.Workbook.FullName } catch {}
			$this.Workbook.Close($true)
			Write-WorkbookCloseLog -ExcelApp $this.App -ProcessId $this.Pid -WorkbookPath $workbookPath
			[System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Workbook) | Out-Null
			$this.Workbook = $null
		}
		# Quit Excel only if we launched it AND we are not keeping the file open
		if (-not $this.ExcelWasOpen -and -not $KeepOpen -and $null -ne $this.App) {
			$excelPid = 0
			try { $excelPid = Get-ExcelProcessId -ExcelApp $this.App } catch { $excelPid = 0 }
			Write-InstanceLog "quitting Excel instance excelPid=$excelPid"
			$this.App.Quit()
			[System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.App) | Out-Null
			$this.App = $null

			# Force release of any remaining COM references so Excel actually
			# exits before this process returns. Without this, Excel lingers and
			# a later command can attach to the dying instance in foreground mode,
			# leaving file locks behind (e.g. `~$` owner files for addin references).
			[System.GC]::Collect()
			[System.GC]::WaitForPendingFinalizers()
			[System.GC]::Collect()
			[System.GC]::WaitForPendingFinalizers()
		} else {
			Write-BridgeLog "quit-skip appWasOpen=$($this.ExcelWasOpen) keepOpen=$KeepOpen appPresent=$($null -ne $this.App)"
		}
	}
}

# -------
# Run
# -------

function Run {
	param(
		[string]$AppName,
		[string]$FilePath,
		[string]$MacroName,
		[bool]$KeepOpen,
		[string[]]$MacroArgValues
	)

	switch ($AppName) {
		"excel" {
			$found = $null
			$targetPath = Get-MacroTargetFile $MacroArgValues
			$lookupPath = if ($FilePath -match '\.(xlam|xla)$' -and $targetPath) {
				$targetPath
			} else {
				$FilePath
			}
			if ($HasRegistry) {
				try {
					$found = Find-OpenWorkbook -Path $lookupPath
				} catch {
					# best-effort; fall back to the normal Excel selection path
				}
			}
			Write-BridgeLog "run-start background=false keepOpen=$KeepOpen file=$FilePath macro=$MacroName lookup=$lookupPath found=$($null -ne $found)"

			$excel = if ($null -ne $found) {
				[Excel]::new($found.App, $found.Workbook)
			} else {
				[Excel]::new()
			}
			$registeredPid = 0
			try {
				# If we created a fresh Excel instance (not attaching to an
				# already-running one), record it in the coordination registry
				# so concurrent agents can distinguish it from a user/rogue
				# session. Best-effort: never break the run over coordination.
				if ($HasRegistry -and -not $excel.ExcelWasOpen) {
					$reason = 'vbapm-run'
					try {
						$registeredPid = Register-ExcelInstance `
							-ExcelApp $excel.App `
							-Owner "terminal-$PID" `
							-Visible $true `
							-Reason $reason `
							-Log (Get-MacroLogDescription $MacroName $MacroArgValues)
					} catch {
						$registeredPid = 0
					}
				}

				$result = $excel.Run($FilePath, $MacroName, $MacroArgValues)
			} catch {
				$result = @{ success = $false; errors = @($_.Exception.Message) } | ConvertTo-Json -Compress
			} finally {
				# Refresh the registry entry's workbook/addin list from the live
				# COM app before teardown, so the deactivated ("inactive") record
				# carries the exact set of workbooks this instance held — letting
				# the end-of-suite assessment trace a lingering instance back to
				# the e2e test that opened them.
				if ($HasRegistry -and $registeredPid -gt 0 -and $null -ne $excel.App) {
					try {
						$log = Get-MacroErrorLog $MacroName $MacroArgValues $result
						Refresh-ExcelInstance -ExcelApp $excel.App -ProcessId $registeredPid -TargetPath (Get-MacroTargetFile $MacroArgValues) -Log $log
					} catch {
						# best-effort
					}
				}

				Write-BridgeLog "dispose-start appWasOpen=$($excel.ExcelWasOpen) isAddin=$($excel.IsAddin) keepOpen=$KeepOpen registeredPid=$registeredPid"
				$excel.Dispose($KeepOpen)
				Write-BridgeLog "dispose-complete appWasOpen=$($excel.ExcelWasOpen) appPresent=$($null -ne $excel.App)"

				# Untrack our instance now that it is torn down.
				if ($HasRegistry -and $registeredPid -gt 0) {
					try {
						Unregister-ExcelInstance -ProcessId $registeredPid
					} catch {
						# best-effort
					}
				}
			}
		}
		default {
			Fail "ERROR #3: Unsupported App `"$AppName`""
		}
	}

	PrintLn $result
}

# -------
# Main
# -------

$ErrorActionPreference = 'Stop'

if (-not $AppName -or -not $File -or -not $Command) {
	Fail "ERROR #1: Invalid Input (appname, file, and macro are required)"
}

if ($MacroArgs.Count -gt 10) {
	Fail "ERROR #2: Invalid Input (only 10 arguments are supported)"
}

# Unescape arguments
$UnescapedArgs = @()
foreach ($arg in $MacroArgs) {
	$UnescapedArgs += Unescape $arg
}

Run $AppName $File $Command $KeepOpen.IsPresent $UnescapedArgs
exit 0
