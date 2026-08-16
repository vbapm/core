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
# The instance is kept visible/hidden according to VBA_BACKGROUND_BUILD (as in
# run.ps1). A workbook opened by a prior request that was NOT "keepOpen" is
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

function Get-ScriptFileName {
	param([string]$Path)
	return [System.IO.Path]::GetFileNameWithoutExtension($Path)
}

function Get-ScriptFileBase {
	param([string]$Path)
	return [System.IO.Path]::GetFileName($Path)
}

function Invoke-ScriptMacro {
	param(
		[object]$ExcelApp,
		[string]$MacroName,
		[string[]]$ArgValues
	)

	$numArgs = if ($ArgValues) { $ArgValues.Count } else { 0 }
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
	return $null
}

function Open-ScriptExcel {
	$Script:BackgroundBuild = $env:VBA_BACKGROUND_BUILD -match '^(1|true|yes)$'

	if (-not $Script:BackgroundBuild) {
		try {
			$Script:App = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
			$Script:AppWasOpen = $true
		} catch {
			# Excel not running; fall through to create a new instance
		}
	}

	if (-not $Script:AppWasOpen) {
		$Script:App = New-Object -ComObject "Excel.Application"
		$Script:App.Visible = if ($Script:BackgroundBuild) { $false } else { $true }
		$Script:App.ScreenUpdating = $false
		$Script:App.DisplayStatusBar = $false
		$Script:App.PrintCommunication = $false
		$Script:App.EnableAnimations = $false
		$Script:App.EnableEvents = $false

		if ($HasRegistry) {
			$reason = if ($Script:BackgroundBuild) { 'e2e' } else { 'vbapm-run' }
			try {
				Register-ExcelInstance -ExcelApp $Script:App -Owner "session#$PID" -Reason $reason | Out-Null
			} catch {
				# best-effort
			}
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

	# Otherwise, if this workbook is already open anywhere, attach to it.
	if ($HasRegistry) {
		try {
			$found = Find-OpenWorkbook -Path $Path
			if ($null -ne $found) {
				$Script:App = $found.App
				$Script:AppWasOpen = $true
				$Script:OpenWorkbook = $found.Workbook
				$Script:WorkbookWasOpen = $true
				return
			}
		} catch {
			# best-effort
		}
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

function Invoke-ScriptRun {
	param(
		[string]$FilePath,
		[string]$MacroName,
		[bool]$KeepOpen,
		[string[]]$ArgValues
	)

	if ($null -eq $Script:App) {
		Open-ScriptExcel
	}

	Ensure-ScriptWorkbook $FilePath

	$result = Invoke-ScriptMacro $Script:App $MacroName $ArgValues

	# Close the workbook unless the caller asked to keep it open (the Application
	# itself is retained across requests so a later command reuses the instance).
	if (-not $KeepOpen -and -not $Script:WorkbookWasOpen -and $null -ne $Script:OpenWorkbook) {
		try {
			$Script:OpenWorkbook.Close($true)
		} catch {
			# best-effort
		}
		$Script:OpenWorkbook = $null
	}

	return $result
}

function Close-ScriptSession {
	# Quit Excel only if we launched it ourselves.
	if (-not $Script:AppWasOpen -and $null -ne $Script:App) {
		try {
			$Script:App.Quit()
			[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Script:App) | Out-Null
		} catch {
			# best-effort
		}
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
	$resolveArgs = @()
	if ($requestJson.args) {
		foreach ($a in $requestJson.args) { $resolveArgs += [string]$a }
	}

	try {
		$result = Invoke-ScriptRun $resolveFile $resolveMacro $resolveKeepOpen $resolveArgs
		$payload = @{ success = $true; result = $result } | ConvertTo-Json -Compress
		Write-ScriptResponse $requestId $payload
	} catch {
		$payload = @{ success = $false; errors = @($_.Exception.Message) } | ConvertTo-Json -Compress
		Write-ScriptResponse $requestId $payload
	}
}
