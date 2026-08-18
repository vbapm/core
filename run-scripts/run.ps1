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

function Fail {
	param([string]$Message)

	# On failure, reveal any background Excel instance we own so a stuck/hung
	# run can be inspected in the UI instead of lingering as a hidden zombie.
	try {
		if ($null -ne $script:ActiveApp -and $script:ActiveIsBackground) {
			$script:ActiveApp.Visible = $true
			PrintErr "ERROR: run failed; making background Excel instance visible for debugging.`n"
		}
	} catch {
		# best-effort; never obscure the original failure
	}

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

# Emit an instance-lifecycle log line, only when VBA_DEBUG_INSTANCES is set.
#
# All lines go to a SINGLE LOG FILE (NOT stderr): stderr is treated as a fatal
# error by the CLI's `toResult` (`if (stderr) { success = false }`), so any
# debug noise on stderr would turn a successful run into a failure. The log path
# is `%TEMP%\Excel-Instances\instances.log` (or VBA_INSTANCE_LOG). Each line is
# timestamped + PID-tagged and includes a live count of EXCEL.EXE processes so
# instance growth/leaks can be tracked by scanning the log.
#
# A "run banner" (clear separator marking the start of a new e2e run) is written
# by the test-side globalSetup before tests start, since `run.ps1` is a
# short-lived process per macro call and can't detect run boundaries.
function Get-InstanceLogPath {
	$path = $env:VBA_INSTANCE_LOG
	if (-not $path) {
		$path = Join-Path $env:TEMP 'Excel-Instances\instances.log'
	}
	return $path
}

function LogInstance {
	param([string]$Message)

	if (-not ($env:VBA_DEBUG_INSTANCES -match '^(1|true|yes)$')) {
		return
	}

	$count = -1
	if (Get-Command Get-RunningExcelInstances -ErrorAction SilentlyContinue) {
		try { $count = @(Get-RunningExcelInstances).Count } catch {}
	}

	$logPath = Get-InstanceLogPath

	$stamp = (Get-Date).ToString('o')
	$line = "[vbapm-instances] $stamp pid=$PID count=$count $Message"

	try {
		Add-Content -LiteralPath $logPath -Value $line -ErrorAction Stop
	} catch {
		# Never let logging itself break a run.
	}
}

# Track the currently-active Excel.Application (script scope) so Fail can reveal
# it if the run aborts. Implemented as a helper function because PowerShell 5.1
# class methods cannot write script-scope variables directly.
function Set-ActiveExcelInstance {
	param([object]$App, [bool]$IsBackground)

	$script:ActiveApp = $App
	$script:ActiveIsBackground = $IsBackground
}

function Clear-ActiveExcelInstance {
	$script:ActiveApp = $null
	$script:ActiveIsBackground = $false
}

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

# Abort if too many EXCEL.EXE processes are already running (safeguard against
# leaked instances piling up). Limit is VBA_MAX_EXCEL_INSTANCES (default 12).
# In background mode, visible (user) instances are excluded from the count so a
# user's normally-open Excel doesn't eat into the automation budget.
# Implemented as a *function* rather than inline class-method code because
# PowerShell 5.1 class methods can't safely reference script-scope variables or
# local variables assigned inside try blocks.
function Assert-ExcelInstanceLimit {
	param([bool]$Background = $false)

	# No-op when the coordination registry module isn't dot-sourced (e.g. an
	# installed CLI without scripts/).
	if (-not (Get-Command Get-RunningExcelInstances -ErrorAction SilentlyContinue)) {
		return
	}

	$maxInstances = 12
	if ($env:VBA_MAX_EXCEL_INSTANCES) {
		$maxInstances = [int]$env:VBA_MAX_EXCEL_INSTANCES
	}

	$instanceCount = 0
	try {
		$running = @(Get-RunningExcelInstances)
		if ($Background) {
			# Only count invisible (automation) instances; ignore visible user sessions.
			$instanceCount = @($running | Where-Object { -not $_.visible }).Count
		} else {
			$instanceCount = $running.Count
		}
	} catch {
		# Best-effort safeguard; never block a legitimate run over a count failure.
		return
	}

	if ($instanceCount -ge $maxInstances) {
		Fail "ERROR #5d: Refusing to open a new Excel instance - $instanceCount EXCEL.EXE process(es) already running (limit $maxInstances). Run scripts/ps/Close-AllInvisibleExcelInstances.ps1 to clean up leaked instances."
	}
}

class Excel {
	hidden [object]$App
	hidden [bool]$BackgroundBuild = $false
	hidden [bool]$ExcelWasOpen = $false
	hidden [object]$Workbook
	hidden [bool]$WorkbookWasOpen = $false
	hidden [bool]$IsAddin = $false
	hidden [int]$Pid = 0

	Excel() {
		$this.OpenExcel()
	}

	[string] Run([string]$FilePath, [string]$MacroName, [string[]]$MacroArgValues) {
		$this.OpenWorkbook($FilePath)

		# Record the just-opened workbook/addin in the coordination registry NOW,
		# while the COM object is still in memory and the workbook is open, and
		# also record the file the macro will operate on (build/export/extract
		# open it inside the macro). Capturing at open time (rather than only at
		# teardown) keeps the workbook list accurate even when the macro closes
		# the workbook before the run returns — which is what lets us attribute a
		# lingering instance to the exact test that opened the workbook.
		if (Get-Command Refresh-ExcelInstance -ErrorAction SilentlyContinue) {
			try {
				Refresh-ExcelInstance -ExcelApp $this.App -ProcessId $this.Pid -TargetPath (Get-MacroTargetFile $MacroArgValues)
			} catch {
				# best-effort; never break a run over coordination bookkeeping
			}
		}

		$result = RunMacro $this.App $MacroName $MacroArgValues

		return $result
	}

	# Attach to a workbook that is already open in a different Excel instance
	# (resolved by the coordination registry helpers).
	[void] Attach([object]$App, [object]$Workbook) {
		$this.App = $App
		$this.ExcelWasOpen = $true
		$this.Workbook = $Workbook
		$this.WorkbookWasOpen = $true
	}

	hidden [void] OpenExcel() {
		# When VBA_BACKGROUND_BUILD is set, always create a new hidden instance
		# instead of attaching to an already-running (visible) Excel process.
		# This prevents the application window from flashing during automated runs.
		$this.BackgroundBuild = $env:VBA_BACKGROUND_BUILD -match '^(1|true|yes)$'

		if (-not $this.BackgroundBuild) {
			try {
				$this.App = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
				$this.ExcelWasOpen = $true
			} catch {
				# Excel not running; fall through to create a new instance
			}
		}

		if (-not $this.ExcelWasOpen) {
			# Safeguard: never spawn a new Excel instance once the machine is
			# already running too many EXCEL.EXE processes. This catches leaked
			# instances from aborted/crashed automated runs before they pile up
			# into the dozens, and aborts (with a clear error) instead of adding
			# one more invisible instance to the pile.
			Assert-ExcelInstanceLimit -Background $this.BackgroundBuild

			try {
				$this.App = New-Object -ComObject "Excel.Application"
				$this.App.Visible = if ($this.BackgroundBuild) { $false } else { $true }
				$this.App.ScreenUpdating = $false
				$this.App.DisplayStatusBar = $false
				$this.App.PrintCommunication = $false
				$this.App.EnableAnimations = $false
				$this.App.EnableEvents = $false
				# Suppress prompts (e.g. "save changes?" on a peer addin loaded
				# via References.AddFromFile) so Close()/Quit() don't block and
				# leave a lingering Excel process holding file locks.
				$this.App.DisplayAlerts = $false
			} catch {
				Fail "ERROR #5: Failed to open Excel - $($_.Exception.Message)"
			}

			$excelPid = 0
			try { $excelPid = Get-ExcelProcessId -ExcelApp $this.App } catch { $excelPid = 0 }
			$this.Pid = $excelPid
			$mode = if ($this.BackgroundBuild) { 'background' } else { 'foreground' }
			LogInstance "created Excel instance excelPid=$excelPid mode=$mode"
			Set-ActiveExcelInstance -App $this.App -IsBackground $this.BackgroundBuild
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
			# Workbooks.Open() can flip Excel back to visible; re-enforce
			# invisible mode if we created this instance for automation.
			# Also apply workbook-level performance flags now that a workbook is open.
			if ($this.BackgroundBuild) {
				$this.App.Visible = $false
				try { $this.App.Calculation = -4135 } catch {}  # xlCalculationManual — requires an open workbook

			}
		} catch {
			Fail "ERROR #6: Failed to open workbook - $($_.Exception.Message)"
		}
	}

	[void] Dispose([bool]$KeepOpen) {
		# An add-in we opened is left open for reuse across runs — but only in
		# FOREGROUND mode, where we're reusing the user's already-running Excel.
		# In BACKGROUND mode each run owns a dedicated hidden instance that is
		# torn down after the run; skipping quit here would leak one EXCEL.EXE
		# per run.
		if ($this.IsAddin -and -not $this.BackgroundBuild) {
			return
		}

		# A file that was open before we started is never closed by us
		$closeWorkbook = -not $this.WorkbookWasOpen -and -not $KeepOpen

		if ($closeWorkbook -and $null -ne $this.Workbook) {
			$this.Workbook.Close($true)
			[System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Workbook) | Out-Null
			$this.Workbook = $null
		}
		# Quit Excel only if we launched it AND we are not keeping the file open
		if (-not $this.ExcelWasOpen -and -not $KeepOpen -and $null -ne $this.App) {
			$excelPid = 0
			try { $excelPid = Get-ExcelProcessId -ExcelApp $this.App } catch { $excelPid = 0 }
			LogInstance "quitting Excel instance excelPid=$excelPid"
			$this.App.Quit()
			[System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.App) | Out-Null
			$this.App = $null
			Clear-ActiveExcelInstance

			# Force release of any remaining COM references so Excel actually
			# exits before this process returns. Without this, Excel lingers and
			# a later command can attach to the dying instance (VBA_BACKGROUND_BUILD=0),
			# leaving file locks behind (e.g. `~$` owner files for addin references).
			[System.GC]::Collect()
			[System.GC]::WaitForPendingFinalizers()
			[System.GC]::Collect()
			[System.GC]::WaitForPendingFinalizers()
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
			$excel = [Excel]::new()
			$registeredPid = 0
			try {
				# If the target workbook is already open in another living Excel
				# instance, attach to that instance + workbook instead of opening
				# a duplicate copy.
				if ($HasRegistry) {
					try {
						$found = Find-OpenWorkbook -Path $FilePath
						if ($null -ne $found) {
							$excel.Attach($found.App, $found.Workbook)
						}
					} catch {
						# best-effort; fall back to opening a fresh copy
					}
				}

				# If we created a fresh Excel instance (not attaching to an
				# already-running one), record it in the coordination registry
				# so concurrent agents can distinguish it from a user/rogue
				# session. Best-effort: never break the run over coordination.
				if ($HasRegistry -and -not $excel.ExcelWasOpen) {
					$reason = if ($excel.BackgroundBuild) { 'e2e' } else { 'vbapm-run' }
					try {
						$registeredPid = Register-ExcelInstance `
							-ExcelApp $excel.App `
							-Owner "terminal-$PID" `
							-Visible (-not $excel.BackgroundBuild) `
							-Reason $reason
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
						Refresh-ExcelInstance -ExcelApp $excel.App -ProcessId $registeredPid -TargetPath (Get-MacroTargetFile $MacroArgValues)
					} catch {
						# best-effort
					}
				}

				$excel.Dispose($KeepOpen)

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
