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
# Excel
# -------

class Excel {
	hidden [object]$App
	hidden [bool]$BackgroundBuild = $false
	hidden [bool]$ExcelWasOpen = $false
	hidden [object]$Workbook
	hidden [bool]$WorkbookWasOpen = $false

	Excel() {
		$this.OpenExcel()
	}

	[string] Run([string]$FilePath, [string]$MacroName, [string[]]$MacroArgValues) {
		$this.OpenWorkbook($FilePath)
		$result = RunMacro $this.App $MacroName $MacroArgValues

		return $result
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
			try {
				$this.App = New-Object -ComObject "Excel.Application"
				$this.App.Visible = if ($this.BackgroundBuild) { $false } else { $true }
				$this.App.ScreenUpdating = $false
				$this.App.DisplayStatusBar = $false
				$this.App.PrintCommunication = $false
				$this.App.EnableAnimations = $false
				$this.App.EnableEvents = $false
			} catch {
				Fail "ERROR #5: Failed to open Excel - $($_.Exception.Message)"
			}
		}
	}

	hidden [void] OpenWorkbook([string]$Path) {
		$fileName = GetFileName $Path
		$fileBase = GetFileBase $Path
		$fullPath = [System.IO.Path]::GetFullPath($Path)

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
		# A file that was open before we started is never closed by us
		$closeWorkbook = -not $this.WorkbookWasOpen -and -not $KeepOpen

		if ($closeWorkbook -and $null -ne $this.Workbook) {
			$this.Workbook.Close($true)
			$this.Workbook = $null
		}
		# Quit Excel only if we launched it AND we are not keeping the file open
		if (-not $this.ExcelWasOpen -and -not $KeepOpen -and $null -ne $this.App) {
			$this.App.Quit()
			[System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.App) | Out-Null
			$this.App = $null
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
			try {
				$result = $excel.Run($FilePath, $MacroName, $MacroArgValues)
			} catch {
				$msg = $_.Exception.Message -replace '"', '\"'
				$result = "{""success"":false,""errors"":[""$msg""]}"
			} finally {
				$excel.Dispose($KeepOpen)
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
