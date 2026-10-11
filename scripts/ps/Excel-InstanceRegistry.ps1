# Excel-InstanceRegistry.ps1
#
# Shared coordinator for tracking Excel.Application instances across concurrent
# agents/processes. It maintains a central registry on disk so that multiple
# automation agents (running on the same machine) don't stomp on each other's
# Excel instances, and so rogue/leftover instances can be identified.
#
# Layout (under %TEMP%\Excel-Instances\):
#   instances.json  - JSON document describing tracked (active) and recently
#                     deactivated (inactive) Excel instances
#   instances.lock  - lock file used to serialize edits to instances.json
#   evolution\     - timestamped instances.json copies when VBA_EVOLUTION=1
#
# This file is a *dot-sourced* library, not a standalone script. It exposes:
#
#   Registry Functions:
#     Get-ExcelInstancesDir        -> path to %TEMP%\Excel-Instances\
#     Get-ExcelInstancesPath       -> full path to instances.json
#     Get-ExcelInstancesLockPath   -> full path to instances.lock
#     Read-ExcelRegistryDocument   -> normalized { active, inactive } document
#     Read-ExcelInstances          -> active instances only (flat array)
#     Read-InactiveExcelInstances  -> recently deactivated instances (flat array)
#     Write-ExcelRegistryDocument  -> write { active, inactive } with lock held
#     Write-ExcelInstances         -> write active entries (preserves inactive)
#     Update-ExcelInstance         -> upsert one active entry (lock held)
#     Refresh-ExcelInstance        -> re-record workbooks/addins from a live app
#     Remove-ExcelInstance         -> deactivate one entry (lock held)
#     Remove-ExcelInstanceById     -> deactivate by stable hash id (lock held)
#
#   Lock Functions:
#     Get-ExcelInstancesLock       -> acquire the lock (blocking, with timeout)
#     Release-ExcelInstancesLock   -> release the lock
#
#   Instance Intrinsics:
#     Get-ExcelProcessId           -> PID of a [ref] to a COM Excel.Application
#     Register-ExcelInstance       -> track a newly created instance
#     Unregister-ExcelInstance     -> deactivate an instance being torn down
#     Get-RunningExcelInstances    -> enumerate live EXCEL.EXE processes w/ PIDs
#     Assert-ExcelInstanceLimit    -> throw if too many EXCEL.EXE are running
#     Get-InstanceLogPath          -> path to the instance-lifecycle debug log
#     Write-InstanceLog            -> append a debug log line (VBA_DEBUG_INSTANCES)
#
# The registry is keyed by process id (Pid). The on-disk document is:
#   {
#     "count":      <int>,   # active + inactive
#     "registered": <int>,   # active entries with a real owner
#     "invisible":  <int>,   # active entries that are hidden
#     "active":     [ ... ], # instances believed to be live
#     "inactive":   [ ... ]  # recently deactivated instances
#   }
#
# Each ACTIVE entry:
#   {
#     "id":        "<random hash>", # stable lower-alnum hash used to close it
#     "pid":        <int>,
#     "owner":      "<string>",   # who created it, e.g. "terminal-12345"
#     "createdAt":  "<ISO-8601>",
#     "visible":    <bool>,       # whether the app window is visible
#     "comReachable": <bool>,     # whether the instance is accessible via COM/ROT
#     "windowTitle": "<string>",  # main window title (may be null)
#     "reason":     "<string>",   # e.g. "e2e", "vbapm-run", "unknown"
#     "intent":     "<string>",   # e.g. "keep-warm" (persists) or "one-shot"
#     "workbooks":  ["<full path>", ...]  # open workbook full paths
#     "logs":       ["<task or error log>", ...]  # instance activity log
#     "workbookReadOnly": {"<normalized path>": <bool>, ...}
#     "addins":     [{"name": "<addin name>", "fullPath": "<path>", "isInstalled": <bool|null>, "isOpen": <bool|null>, "readOnly": <bool>}, ...]
#     "ghostWorkbooks": [{"name": "<VBE project>", "file": "<path>", ...}, ...]
#     "ghostDetection": {"available": <bool>, "ghostCount": <int>, ...}
#   }
#
# An INACTIVE entry is a full active entry that was deactivated instead of being
# deleted, plus:
#   {
#     "deactivatedAt":     "<ISO-8601>",  # when it was moved to inactive
#     "deactivateReason":  "<string>"     # e.g. "unregistered", "dead-process"
#   }
#
# Keeping deactivated entries (rather than deleting them) lets the end-of-suite
# assessment (Assess-ExcelInstances.ps1) detect a "zombie": an instance whose
# registry entry was deactivated but whose EXCEL.EXE process is still alive —
# and attribute it to the workbooks (and therefore the e2e test) that were open
# on it when it was deactivated.
#
# All access to instances.json is serialized through the lock file so that
# concurrent agents never read a half-written JSON document.

$ErrorActionPreference = 'Stop'

# -------
# Paths
# -------

function Get-ExcelInstancesDir {
    return Join-Path $env:TEMP 'Excel-Instances'
}

function Get-ExcelInstancesPath {
    return Join-Path (Get-ExcelInstancesDir) 'instances.json'
}

function Get-ExcelInstancesLockPath {
    return Join-Path (Get-ExcelInstancesDir) 'instances.lock'
}

function Get-ExcelInstancesEvolutionDir {
    return Join-Path (Get-ExcelInstancesDir) 'evolution'
}

function Test-ExcelRegistryEvolution {
    return [string]$env:VBA_EVOLUTION -match '^(1|true|yes)$'
}

function Get-ExcelInstancesEvolutionRunDir {
    $evolutionDir = Get-ExcelInstancesEvolutionDir
    $marker = Join-Path $evolutionDir 'current-run.txt'

    if ($env:VBA_EVOLUTION_RUN_ID) {
        New-Item -ItemType Directory -Path $evolutionDir -Force | Out-Null
        $runDir = Join-Path $evolutionDir $env:VBA_EVOLUTION_RUN_ID
        if (-not (Test-Path -LiteralPath $runDir -PathType Container)) {
            New-Item -ItemType Directory -Path $runDir -Force | Out-Null
        }
        [System.IO.File]::WriteAllText($marker, $runDir, [System.Text.UTF8Encoding]::new($false))
        return $runDir
    }

    if (Test-Path -LiteralPath $marker -PathType Leaf) {
        $runDir = (Get-Content -LiteralPath $marker -Raw -Encoding UTF8).Trim()
        if ($runDir -and (Test-Path -LiteralPath $runDir -PathType Container)) {
            return $runDir
        }
    }

    New-Item -ItemType Directory -Path $evolutionDir -Force | Out-Null
    $stamp = [DateTime]::UtcNow.ToString("yyyy-MM-dd--HH'h'mm'm'ss's'fff")
    $runDir = Join-Path $evolutionDir $stamp
    $suffix = 1
    while (Test-Path -LiteralPath $runDir) {
        $runDir = Join-Path $evolutionDir "$stamp-$suffix"
        $suffix++
    }
    New-Item -ItemType Directory -Path $runDir -Force | Out-Null
    [System.IO.File]::WriteAllText($marker, $runDir, [System.Text.UTF8Encoding]::new($false))
    return $runDir
}

function Save-ExcelRegistryEvolutionSnapshot {
    param(
        [Parameter(Mandatory)][string]$Json,
        [Parameter(Mandatory)][string]$JsonPath
    )

    if (-not (Test-ExcelRegistryEvolution)) { return }

    $evolutionDir = Get-ExcelInstancesEvolutionRunDir

    $stamp = [DateTime]::UtcNow.ToString("yyyy-MM-dd--HH'h'mm'm'ss's'fff")
    $snapshotPath = Join-Path $evolutionDir "instances-$stamp.json"
    $suffix = 1
    while (Test-Path -LiteralPath $snapshotPath) {
        $snapshotPath = Join-Path $evolutionDir "instances-$stamp-$suffix.json"
        $suffix++
    }

    [System.IO.File]::WriteAllText($snapshotPath, $Json, [System.Text.UTF8Encoding]::new($false))
}

# -------
# Lock
# -------

<#
.SYNOPSIS
Acquire the registry lock, blocking until it is available (or timeout).

Uses an atomic "create-if-absent" on the lock file. The lock file stores the
owner (PID + timestamp) for diagnostics. Stale locks (older than the timeout)
belonging to a dead process are stolen so a crashed agent can't wedge everyone.
#>
function Get-ExcelInstancesLock {
    param(
        [int]$TimeoutSeconds = 60,
        [int]$PollMilliseconds = 100
    )

    if ($null -ne $script:ExcelInstancesLockStream) {
        $script:ExcelInstancesLockDepth++
        return $true
    }

    $dir = Get-ExcelInstancesDir
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $lockPath = Get-ExcelInstancesLockPath
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $ownerId = "$PID@$([Environment]::MachineName)"

    while ($true) {
        try {
            # Atomic creation; throws if the file already exists.
            $stream = New-Object System.IO.FileStream(
                $lockPath,
                [System.IO.FileMode]::CreateNew,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::None,
                256,
                [System.IO.FileOptions]::DeleteOnClose
            )
            $writer = New-Object System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8, 256, $true)
            $writer.Write("$ownerId $((Get-Date).ToString('o'))")
            $writer.Flush()
            $writer.Dispose()
            $script:ExcelInstancesLockStream = $stream
            $script:ExcelInstancesLockDepth = 1
            return $true
        } catch [System.IO.IOException] {
            # Lock file already exists. Check whether it is stale.
            $stale = $false
            if (Test-Path -LiteralPath $lockPath) {
                $lastWrite = (Get-Item -LiteralPath $lockPath).LastWriteTime
                if ((Get-Date) - $lastWrite -gt [TimeSpan]::FromSeconds($TimeoutSeconds)) {
                    $stale = $true
                }
            }

            if ($stale) {
                # Best-effort steal of a stale lock. If this fails (lost race),
                # fall through and keep waiting.
                try {
                    Remove-Item -LiteralPath $lockPath -Force -ErrorAction Stop
                    continue
                } catch {
                    # Another process stole it first; keep waiting.
                }
            }

            if ((Get-Date) -ge $deadline) {
                throw "Timed out after $TimeoutSeconds seconds waiting for the Excel instance registry lock ($lockPath)."
            }
            Start-Sleep -Milliseconds $PollMilliseconds
        }
    }
}

<#
.SYNOPSIS
Release a previously acquired registry lock.
#>
function Release-ExcelInstancesLock {
    if ($null -ne $script:ExcelInstancesLockStream) {
        $script:ExcelInstancesLockDepth--
        if ($script:ExcelInstancesLockDepth -gt 0) {
            return
        }

        try { $script:ExcelInstancesLockStream.Dispose() } catch {}
        $script:ExcelInstancesLockStream = $null
        $script:ExcelInstancesLockDepth = 0
    }

    # DeleteOnClose removes our lock atomically when the stream is disposed.
    # Do not remove the path here: another process may acquire a new lock in
    # the interval between disposal and an explicit Remove-Item call.
}

# -------
# Registry read/write
# -------

<#
.SYNOPSIS
Read the full registry document, normalized to { active, inactive }. Creates an
empty registry (and directory) if none exists. Safe to call without the lock.

Normalizes legacy on-disk shapes (a bare array, or the older
`{ count, registered, invisible, instances }` object) to the new
{ active, inactive } document so old files still load.
#>
function Read-ExcelRegistryDocument {
    $dir = Get-ExcelInstancesDir
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $jsonPath = Get-ExcelInstancesPath
    if (-not (Test-Path -LiteralPath $jsonPath)) {
        return [ordered]@{ active = @(); inactive = @() }
    }

    try {
        $raw = Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8 -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return [ordered]@{ active = @(); inactive = @() }
        }
        $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
        if ($null -eq $parsed) {
            return [ordered]@{ active = @(); inactive = @() }
        }

        $active = @()
        $inactive = @()

        if ($parsed -is [array]) {
            # Legacy: bare array of instances.
            $active = @($parsed)
        } elseif ($parsed.PSObject.Properties.Name -contains 'active') {
            $active = @($parsed.active)
            if ($parsed.PSObject.Properties.Name -contains 'inactive') {
                $inactive = @($parsed.inactive)
            }
        } elseif ($parsed.PSObject.Properties.Name -contains 'instances') {
            # Older schema: { count, registered, invisible, instances }.
            $active = @($parsed.instances)
        } else {
            $active = @($parsed)
        }

        return [ordered]@{ active = $active; inactive = $inactive }
    } catch {
        Write-Warning "Failed to read Excel instance registry: $($_.Exception.Message)"
        return [ordered]@{ active = @(); inactive = @() }
    }
}

<#
.SYNOPSIS
Read the active instances only, as a flat array (backwards-compatible with
callers that predate the active/inactive split).
#>
function Read-ExcelInstances {
    return @((Read-ExcelRegistryDocument).active)
}

<#
.SYNOPSIS
Read the recently deactivated instances as a flat array.
#>
function Read-InactiveExcelInstances {
    return @((Read-ExcelRegistryDocument).inactive)
}

<#
.SYNOPSIS
Clear the inactive (recently deactivated) instance list. Called at the start of
an e2e run so stale entries from previous runs don't pollute the end-of-suite
assessment. Holds the lock for the duration.
#>
function Clear-InactiveExcelInstances {
    Get-ExcelInstancesLock | Out-Null
    try {
        $doc = Read-ExcelRegistryDocument
        Write-ExcelRegistryDocument -Active $doc.active -Inactive @()
    } finally {
        Release-ExcelInstancesLock
    }
}

function Clear-ActiveE2EExcelInstanceLogs {
    Get-ExcelInstancesLock | Out-Null
    try {
        $doc = Read-ExcelRegistryDocument
        $changed = $false
        foreach ($entry in @($doc.active)) {
            $isHiddenE2E = -not [bool]$entry.visible -and (
                [string]$entry.reason -eq 'e2e' -or
                [string]$entry.owner -like 'session#*'
            )
            if ($isHiddenE2E -and @($entry.logs).Count -gt 0) {
                $entry.logs = @()
                $changed = $true
            }
        }
        if ($changed) {
            Write-ExcelRegistryDocument -Active $doc.active -Inactive $doc.inactive
        }
    } finally {
        Release-ExcelInstancesLock
    }
}

function ConvertTo-ExcelRegistryEventEntry {
    param([object]$Entry)

    return [ordered]@{
        id               = [string]$Entry.id
        pid              = [int]$Entry.pid
        owner            = [string]$Entry.owner
        reason           = [string]$Entry.reason
        visible          = [bool]$Entry.visible
        comReachable     = [bool]$Entry.comReachable
        workbooks        = @($Entry.workbooks)
        workbookReadOnly = $Entry.workbookReadOnly
        addins           = @($Entry.addins)
        logs             = @($Entry.logs)
        ghostWorkbooks   = @($Entry.ghostWorkbooks)
        ghostDetection   = $Entry.ghostDetection
        deactivatedAt    = [string]$Entry.deactivatedAt
        deactivateReason = [string]$Entry.deactivateReason
    }
}

function Get-ExcelRegistryEventSignature {
    param(
        [object[]]$Active,
        [object[]]$Inactive
    )

    $eventDocument = [ordered]@{
        active   = @($Active | ForEach-Object { ConvertTo-ExcelRegistryEventEntry $_ })
        inactive = @($Inactive | ForEach-Object { ConvertTo-ExcelRegistryEventEntry $_ })
    }
    return ($eventDocument | ConvertTo-Json -Depth 20 -Compress)
}

<#
.SYNOPSIS
Write the full registry document (active + inactive) as JSON. Callers should
hold the lock.
#>
function Write-ExcelRegistryDocument {
    param(
        [object[]]$Active,
        [object[]]$Inactive = @(),
        [switch]$ForceEvolutionSnapshot
    )

    $dir = Get-ExcelInstancesDir
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $jsonPath = Get-ExcelInstancesPath

    $activeList = @($Active)
    $inactiveList = @($Inactive)

    # Assign a stable, human-friendly number to every entry. Existing numbers
    # are preserved; missing or duplicate numbers receive the next available
    # value. Active entries are assigned before inactive history.
    $usedNumbers = @{}
    $nextNumber = 1
    foreach ($entry in @($activeList + $inactiveList)) {
        $entryNumber = 0
        try { $entryNumber = [int]$entry.number } catch {}
        if ($entryNumber -gt 0 -and -not $usedNumbers.ContainsKey($entryNumber)) {
            $usedNumbers[$entryNumber] = $true
            if ($entryNumber -ge $nextNumber) { $nextNumber = $entryNumber + 1 }
        } else {
            while ($usedNumbers.ContainsKey($nextNumber)) { $nextNumber++ }
            if ($null -eq $entry.PSObject.Properties['number']) {
                $entry | Add-Member -MemberType NoteProperty -Name number -Value $nextNumber
            } else {
                $entry.number = $nextNumber
            }
            $usedNumbers[$nextNumber] = $true
            $nextNumber++
        }
    }

    # "registered" = entries that carry a real owner (explicitly tracked by an
    # agent), as opposed to `rogue`/`unknown` entries discovered only via the
    # live process list.
    $registered = @($activeList | Where-Object {
        $o = [string]$_.owner
        $o -and $o -notmatch '^(rogue|unknown|unknown-owner)$'
    }).Count

    # "invisible" = automation-created (hidden) instances: visible is false or
    # unset. Useful for spotting leaked background instances at a glance.
    $invisible = @($activeList | Where-Object { -not [bool]$_.visible }).Count

    $doc = [ordered]@{
        count      = $activeList.Count + $inactiveList.Count
        registered = $registered
        invisible  = $invisible
        active     = $activeList
        inactive   = $inactiveList
    }

    $json = $doc | ConvertTo-Json -Depth 6
    $previousJson = $null
    if (Test-Path -LiteralPath $jsonPath) {
        try { $previousJson = [System.IO.File]::ReadAllText($jsonPath, [System.Text.UTF8Encoding]::new($false)) } catch {}
    }

    $eventChanged = $true
    if ($null -ne $previousJson) {
        try {
            $previousDocument = $previousJson | ConvertFrom-Json -ErrorAction Stop
            $eventChanged = (Get-ExcelRegistryEventSignature `
                -Active @($previousDocument.active) `
                -Inactive @($previousDocument.inactive)) -ne (Get-ExcelRegistryEventSignature `
                -Active $activeList `
                -Inactive $inactiveList)
        } catch {
            $eventChanged = $true
        }
    }

    if ($previousJson -ne $json) {
        [System.IO.File]::WriteAllText($jsonPath, $json, [System.Text.UTF8Encoding]::new($false))
        if ($eventChanged -or $ForceEvolutionSnapshot) {
            Save-ExcelRegistryEvolutionSnapshot -Json $json -JsonPath $jsonPath
        }
    }
}

<#
.SYNOPSIS
Write the active instances (overwriting). Callers should hold the lock. Existing
inactive entries are preserved so a partial write doesn't lose deactivated
instances; use Write-ExcelRegistryDocument to change both lists at once.
#>
function Write-ExcelInstances {
    param([object[]]$Instances)

    # Callers hold the lock, so reading the current inactive list here is safe.
    $existingInactive = @((Read-ExcelRegistryDocument).inactive)
    Write-ExcelRegistryDocument -Active $Instances -Inactive $existingInactive
}

# -------
# Instance operations (lock-guarded)
# -------

<#
.SYNOPSIS
Return the full paths of all open workbooks in a COM Excel.Application [ref].
Unsafe/new (unsaved) workbooks have no path and are reported by name in
brackets.
#>
function Get-ExcelWorkbooks {
    param([Parameter(Mandatory)][object]$ExcelApp)

    $workbooks = @()
    try {
        foreach ($wb in $ExcelApp.Workbooks) {
            try {
                $fullName = $wb.FullName
                $name = $wb.Name
            } catch {
                continue
            }

            if ($fullName -and $fullName -ne $name) {
                # Normalize to forward slashes for cross-platform consistency.
                $workbooks += ($fullName -replace '\\', '/')
            } else {
                # Unsaved/new workbook — no FullName; fall back to its name.
                $workbooks += "[$name]"
            }
        }
    } catch {
        # Workbooks collection unavailable; return whatever we have.
    }
    return $workbooks
}

<#
.SYNOPSIS
Return read-only state keyed by the normalized workbook value used in the
instance registry.
#>
function Get-ExcelWorkbookReadOnly {
    param([Parameter(Mandatory)][object]$ExcelApp)

    $readOnlyByWorkbook = @{}
    try {
        foreach ($workbook in $ExcelApp.Workbooks) {
            try {
                $fullName = [string]$workbook.FullName
                $name = [string]$workbook.Name
                $registryValue = if ($fullName -and $fullName -ne $name) {
                    ($fullName -replace '\\', '/')
                } else {
                    "[$name]"
                }
                $key = $registryValue.ToLowerInvariant()
                $readOnlyByWorkbook[$key] = [bool]$workbook.ReadOnly
            } catch {}
        }
    } catch {}
    return $readOnlyByWorkbook
}

<#
.SYNOPSIS
Return the names (and load status) of add-ins known to a COM Excel.Application
[ref]. Each add-in reports two independent flags:
  `isInstalled` — whether it is registered/checked in the Add-Ins dialog
                  (`AddIn.Installed`).
  `isOpen`      — whether it is actually loaded into this Excel process
                  (`AddIn.IsOpen`). An add-in can be `Installed=True` without
                  being loaded (common on automation-created instances), so
                  `isOpen` must NOT be derived from `Installed`. If a flag
                  cannot be read it is reported as `null` ("unknown") rather
                  than fabricating a state.
Names are unique per add-in and normalized to forward slashes for any path-like
names. `fullPath` contains the add-in's COM `FullName` when Excel provides it.
`readOnly` comes from the matching open workbook.
#>
function Get-ExcelAddins {
    param([Parameter(Mandatory)][object]$ExcelApp)

    $addins = @()
    $readOnlyByPath = @{}
    try {
        foreach ($workbook in $ExcelApp.Workbooks) {
            try {
                $workbookPath = [string]$workbook.FullName
                if ($workbookPath) {
                    $normalizedPath = ($workbookPath -replace '\\', '/').ToLowerInvariant()
                    $readOnlyByPath[$normalizedPath] = [bool]$workbook.ReadOnly
                }
            } catch {}
        }
    } catch {}

    try {
        foreach ($addin in $ExcelApp.AddIns) {
            $name = $null
            try { $name = [string]$addin.Name } catch { $name = $null }
            if (-not $name) { continue }

            # isInstalled = registered/checked in the Add-Ins dialog (Installed).
            # isOpen = actually loaded into this Excel process (IsOpen). Read
            # them independently; an add-in can be Installed=True yet not loaded
            # (common on automation-created instances), so isOpen must not be
            # derived from Installed. If a read fails, report $null ("unknown")
            # rather than fabricating a state or throwing into a swallowing
            # caller.
            $isInstalled = $null
            try { $isInstalled = [bool]$addin.Installed } catch { $isInstalled = $null }

            $isOpen = $null
            try { $isOpen = [bool]$addin.IsOpen } catch { $isOpen = $null }

            $fullPath = $null
            try {
                if ($addin.FullName) {
                    $fullPath = ([string]$addin.FullName -replace '\\', '/')
                }
            } catch {}

            $readOnly = $false
            if ($fullPath) {
                $normalizedPath = $fullPath.ToLowerInvariant()
                if ($readOnlyByPath.ContainsKey($normalizedPath)) {
                    $readOnly = [bool]$readOnlyByPath[$normalizedPath]
                }
            }

            $addins += [pscustomobject]@{
                name         = ($name -replace '\\', '/')
                fullPath     = $fullPath
                isInstalled  = $isInstalled
                isOpen       = $isOpen
                readOnly     = $readOnly
            }
        }
    } catch {
        # AddIns collection unavailable; return whatever we have.
    }
    return @($addins)
}

<#
.SYNOPSIS
Ensure a COM Excel.Application [ref] has the given add-in installed and loaded.
Reuses an already-loaded add-in, otherwise adds + installs it. Returns the add-in
COM object (or $null on failure).

This is the mechanism for leaving `vbapm.xlam` open across runs: once installed,
the add-in stays in the Application's AddIns collection for the life of the Excel
process, so subsequent macro runs reuse it instead of reopening it.
#>
function Ensure-ExcelAddin {
    param(
        [Parameter(Mandatory)][object]$ExcelApp,
        [Parameter(Mandatory)][string]$AddinPath,
        [string]$AddinName = ''
    )

    $fullPath = [System.IO.Path]::GetFullPath($AddinPath)
    $fileName = [System.IO.Path]::GetFileName($fullPath)
    $name = if ($AddinName) { $AddinName } else { [System.IO.Path]::GetFileNameWithoutExtension($fullPath) }

    # 1) Already installed/loaded via the AddIns collection?
    try {
        $addin = $ExcelApp.AddIns($fileName)
        if ($addin -and ([bool]$addin.Installed)) {
            return $addin
        }
    } catch {
        # Not found by that name; fall through to add it.
    }

    # 2) Already open as a workbook (some flows open the .xlam directly)?
    try {
        foreach ($wb in $ExcelApp.Workbooks) {
            if ($wb.FullName -eq $fullPath) {
                return $wb
            }
        }
    } catch {
        # ignore
    }

    # 3) Add + install it as a proper add-in.
    try {
        $addin = $ExcelApp.AddIns.Add($fullPath, $true)
        $addin.Installed = $true
        return $addin
    } catch {
        Write-Warning "Failed to ensure add-in '$fileName' is loaded: $($_.Exception.Message)"
        return $null
    }
}

<#
.SYNOPSIS
Generate a random lowercase-alphanumeric id (hash) for an Excel instance. Used
so agents can refer to (and close) a specific instance independent of its PID.
#>
function New-ExcelInstanceId {
    param([int]$Length = 8)

    $chars = 'abcdefghijklmnopqrstuvwxyz0123456789'
    $sb = New-Object System.Text.StringBuilder
    $random = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $bytes = New-Object byte[] $Length
    $random.GetBytes($bytes)
    foreach ($b in $bytes) {
        [void]$sb.Append($chars[$b % $chars.Length])
    }
    return $sb.ToString()
}

<#
.SYNOPSIS
Upsert an instance entry keyed by its PID. Holds the lock for the duration.
The `id` hash is generated on first creation and preserved on subsequent
updates.
#>
function Update-ExcelInstance {
    param(
        [Parameter(Mandatory)][int]$InstanceId,
        [Parameter(Mandatory)][string]$Owner,
        [bool]$Visible = $false,
        [string]$Reason = 'unknown',
        [string[]]$Workbooks = @(),
        [string[]]$Logs = $null,
        [object]$WorkbookReadOnly = $null,
        [object[]]$Addins = @(),
        [object[]]$GhostWorkbooks = @(),
        [object]$GhostDetection = $null,
        [string]$WindowTitle = '',
        [bool]$ComReachable = $true,
        # Whether this instance is intended to be kept alive (warm) beyond the
        # operation that created it. 'keep-warm' = a long-lived/persistent session
        # instance; 'one-shot' = created for a single operation then torn down.
        [string]$Intent = ''
    )

    Get-ExcelInstancesLock | Out-Null
    try {
        $doc = Read-ExcelRegistryDocument
        $instances = @($doc.active)

        # Preserve an existing id hash if this PID is already tracked (active or
        # previously deactivated), so the id stays stable across updates and
        # re-activations.
        $existingId = $null
        foreach ($inst in $instances) {
            if ([int]$inst.pid -eq $InstanceId -and $inst.id) {
                $existingId = [string]$inst.id
                break
            }
        }
        if (-not $existingId) {
            foreach ($inst in @($doc.inactive)) {
                if ([int]$inst.pid -eq $InstanceId -and $inst.id) {
                    $existingId = [string]$inst.id
                    break
                }
            }
        }
        if (-not $existingId) {
            $existingId = New-ExcelInstanceId
        }

        $hasNewLog = $null -ne $Logs -and @($Logs | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0
        $existingEntry = @($instances | Where-Object { [int]$_.pid -eq $InstanceId })[0]
        if ($null -eq $Logs) {
            $Logs = if ($existingEntry -and $null -ne $existingEntry.logs) { @($existingEntry.logs) } else { @() }
        }
        $Logs = @($Logs | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
        if ($null -eq $WorkbookReadOnly) {
            $WorkbookReadOnly = if ($existingEntry -and $null -ne $existingEntry.workbookReadOnly) {
                $existingEntry.workbookReadOnly
            } else {
                @{}
            }
        }
        # Preserve an existing intent unless the caller explicitly overrides it,
        # so refreshes/syncs never downgrade a warm instance to one-shot.
        if ([string]::IsNullOrWhiteSpace($Intent)) {
            $Intent = if ($existingEntry -and $existingEntry.intent) { [string]$existingEntry.intent } else { 'one-shot' }
        }
		$warmOwner = if ($existingEntry -and $existingEntry.PSObject.Properties['warmOwner']) { [string]$existingEntry.warmOwner } else { '' }
		$warmAvailable = if ($existingEntry -and $existingEntry.PSObject.Properties['warmAvailable']) { [bool]$existingEntry.warmAvailable } else { $false }

        $entry = @{
            id           = $existingId
            number       = $null
            pid          = $InstanceId
            owner        = $Owner
            visible      = $Visible
            reason       = $Reason
            intent       = $Intent
			warmOwner    = $warmOwner
			warmAvailable = $warmAvailable
            comReachable = [bool]$ComReachable
            windowTitle  = if ([string]::IsNullOrEmpty($WindowTitle)) { $null } else { $WindowTitle }
            createdAt    = (Get-Date).ToString('o')
            workbooks    = if ($null -eq $Workbooks) { $null } else { @($Workbooks) }
            logs         = @($Logs)
            workbookReadOnly = $WorkbookReadOnly
            addins       = if ($null -eq $Addins) { $null } else { @($Addins) }
            ghostWorkbooks = @($GhostWorkbooks)
            ghostDetection = $GhostDetection
        }

        $found = $false
        for ($i = 0; $i -lt $instances.Count; $i++) {
            if ([int]$instances[$i].pid -eq $InstanceId) {
                $instances[$i] = [pscustomobject]$entry
                $found = $true
                break
            }
        }
        if (-not $found) {
            $instances += [pscustomobject]$entry
        }

        # This PID is active again — drop any stale inactive entry for it.
        $inactive = @($doc.inactive | Where-Object { [int]$_.pid -ne $InstanceId })

        Write-ExcelRegistryDocument -Active $instances -Inactive $inactive -ForceEvolutionSnapshot:$hasNewLog
    } finally {
        Release-ExcelInstancesLock
    }
}

function Set-ExcelInstanceWarmState {
    param(
        [Parameter(Mandatory)][int]$ProcessId,
        [string]$WarmOwner = '',
        [bool]$WarmAvailable = $false
    )

    Get-ExcelInstancesLock | Out-Null
    try {
        $doc = Read-ExcelRegistryDocument
        $active = @($doc.active)
        foreach ($entry in $active) {
            if ([int]$entry.pid -ne $ProcessId) { continue }
            if ($null -eq $entry.PSObject.Properties['warmOwner']) {
                $entry | Add-Member -MemberType NoteProperty -Name warmOwner -Value $WarmOwner
            } else { $entry.warmOwner = $WarmOwner }
            if ($null -eq $entry.PSObject.Properties['warmAvailable']) {
                $entry | Add-Member -MemberType NoteProperty -Name warmAvailable -Value ([bool]$WarmAvailable)
            } else { $entry.warmAvailable = [bool]$WarmAvailable }
            Write-ExcelRegistryDocument -Active $active -Inactive $doc.inactive
            return $true
        }
        return $false
    } finally {
        Release-ExcelInstancesLock
    }
}

<#
.SYNOPSIS
Annotate an active entry for the inactive list: copy it and stamp when/why it
was deactivated.
#>
function ConvertTo-InactiveEntry {
    param(
        [object]$Entry,
        [string]$DeactivateReason
    )

    $deactivated = $Entry | Select-Object *
    $deactivated | Add-Member -NotePropertyName deactivatedAt -NotePropertyValue ((Get-Date).ToString('o')) -Force
    $deactivated | Add-Member -NotePropertyName deactivateReason -NotePropertyValue $DeactivateReason -Force
    return $deactivated
}

<#
.SYNOPSIS
Deactivate an instance entry by PID: move it from `active` to `inactive` (rather
than deleting it), so a later assessment can tell whether the process actually
exited. Holds the lock for the duration.
#>
function Remove-ExcelInstance {
    param(
        [Parameter(Mandatory)][int]$InstanceId,
        [string]$DeactivateReason = 'unregistered'
    )

    Get-ExcelInstancesLock | Out-Null
    try {
        $doc = Read-ExcelRegistryDocument
        $active = @($doc.active)
        $inactive = @($doc.inactive)

        $removed = @($active | Where-Object { [int]$_.pid -eq $InstanceId })
        $kept = @($active | Where-Object { [int]$_.pid -ne $InstanceId })

        foreach ($entry in $removed) {
            $inactive += (ConvertTo-InactiveEntry -Entry $entry -DeactivateReason $DeactivateReason)
        }

        # Keep the inactive list bounded to the most recent entries.
        $maxInactive = 100
        if ($inactive.Count -gt $maxInactive) {
            $inactive = @($inactive | Select-Object -Last $maxInactive)
        }

        Write-ExcelRegistryDocument -Active $kept -Inactive $inactive
    } finally {
        Release-ExcelInstancesLock
    }
}

<#
.SYNOPSIS
Look up a registry entry by its stable hash id. Returns $null if not found.
#>
function Find-ExcelInstanceById {
    param([Parameter(Mandatory)][string]$Id)

    $instances = @(Read-ExcelInstances)
    foreach ($inst in $instances) {
        if ($inst.id -and [string]$inst.id -eq $Id) {
            return $inst
        }
    }
    return $null
}

<#
.SYNOPSIS
Look up an active registry entry by its human-friendly number.
#>
function Find-ExcelInstanceByNumber {
    param([Parameter(Mandatory)][int]$Number)

    $instances = @(Read-ExcelInstances)
    foreach ($inst in $instances) {
        if ($inst.number -and [int]$inst.number -eq $Number) {
            return $inst
        }
    }
    return $null
}

<#
.SYNOPSIS
Deactivate a registry entry by its stable hash id (move to `inactive`). Holds
the lock for the duration.
#>
function Remove-ExcelInstanceById {
    param(
        [Parameter(Mandatory)][string]$Id,
        [string]$DeactivateReason = 'closed-by-id'
    )

    Get-ExcelInstancesLock | Out-Null
    try {
        $doc = Read-ExcelRegistryDocument
        $active = @($doc.active)
        $inactive = @($doc.inactive)

        $removed = @($active | Where-Object { $_.id -and [string]$_.id -eq $Id })
        $kept = @($active | Where-Object { -not ($_.id -and [string]$_.id -eq $Id) })

        foreach ($entry in $removed) {
            $inactive += (ConvertTo-InactiveEntry -Entry $entry -DeactivateReason $DeactivateReason)
        }

        $maxInactive = 100
        if ($inactive.Count -gt $maxInactive) {
            $inactive = @($inactive | Select-Object -Last $maxInactive)
        }

        Write-ExcelRegistryDocument -Active $kept -Inactive $inactive
    } finally {
        Release-ExcelInstancesLock
    }
}

<#
.SYNOPSIS
Close (Quit) a running Excel process by its stable hash id, then remove it from
the registry. Returns $true on success, $false when no matching instance exists.
#>
function Close-ExcelInstanceById {
    param([Parameter(Mandatory)][string]$Id)

    $entry = Find-ExcelInstanceById -Id $Id
    if (-not $entry) {
        return $false
    }

    $pidToClose = [int]$entry.pid

    # Best-effort: locate the process and ask it to quit. If the process is not
    # reachable via COM (not in ROT), fall back to Stop-Process.
    $quit = $false
    try {
        $proc = Get-Process -Id $pidToClose -ErrorAction Stop
        if ($proc.ProcessName -eq 'EXCEL') {
            try {
                $proc.CloseMainWindow() | Out-Null
                $quit = $true
            } catch {
                # fall through to Stop-Process
            }
            if (-not $quit -or -not $proc.HasExited) {
                try { Stop-Process -Id $pidToClose -Force -ErrorAction Stop; $quit = $true } catch { $quit = $false }
            }
        } else {
            $quit = $false
        }
    } catch {
        $quit = $false
    }

    # Remove from registry regardless of whether the process could be killed.
    Remove-ExcelInstanceById -Id $Id

    return $quit
}

# -------
# Live-instance intrinsics
# -------

<#
.SYNOPSIS
Extract the process id from a COM Excel.Application [ref]. `$App.Hwnd` is the
application's top-level window handle; walking up the window owner chain is
unreliable, so we use the Running Object Table (ROT) matching on Hwnd, with a
fallback to enumerating EXCEL.EXE processes whose main window handle matches.
#>
function Get-ExcelProcessId {
    param([Parameter(Mandatory)][object]$ExcelApp)

    # Hwnd is exposed by Excel.Application on Windows. Resolve it to a PID
    # using GetWindowThreadProcessId (reliable across instances, including
    # hidden ones whose MainWindowHandle is 0).
    $hwnd = $null
    try {
        $hwnd = [int64]$ExcelApp.Hwnd
    } catch {
        $hwnd = $null
    }

    $InstanceClass = 'ExcelInstanceWin32'
    if (-not ($InstanceClass -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class ExcelInstanceWin32 {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
}
"@
    }

    if ($hwnd -gt 0) {
        $pidValue = [uint32]0
        try {
            [ExcelInstanceWin32]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$pidValue) | Out-Null
            if ($pidValue -gt 0) {
                return [int]$pidValue
            }
        } catch {
            # fall through
        }

        # Fallback: match MainWindowHandle.
        $procs = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
        foreach ($p in $procs) {
            try {
                if ([int64]$p.MainWindowHandle -eq $hwnd) {
                    return [int]$p.Id
                }
            } catch {
                # ignore
            }
        }
    }

    # Last resort: if only one EXCEL.EXE exists, assume it is ours.
    $procs = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
    if ($procs.Count -eq 1) {
        return [int]$procs[0].Id
    }

    throw "Could not determine process id for the Excel.Application instance."
}

<#
.SYNOPSIS
Register a newly created Excel.Application instance in the registry.
#>
function Register-ExcelInstance {
    param(
        [Parameter(Mandatory)][object]$ExcelApp,
        [string]$Owner,
        [bool]$Visible = $false,
        [string]$Reason = 'unknown',
        [string]$Log = '',
        [string]$Intent = 'one-shot'
    )

    if (-not $Owner) {
        # Trace ownership back to the terminal/process that invoked this
        # function. Since this library is dot-sourced into the caller, $PID is
        # the calling terminal's process id.
        $Owner = "terminal-$PID"
    }

    try {
        $instancePid = Get-ExcelProcessId -ExcelApp $ExcelApp

        # Capture visibility from the live app if not explicitly provided.
        $isVisible = $Visible
        try {
            $isVisible = [bool]$ExcelApp.Visible
        } catch {
            # keep the caller-supplied value
        }

        # Capture the open-workbook list from the live app.
        $openWorkbooks = @(Get-ExcelWorkbooks -ExcelApp $ExcelApp)

        # Capture the loaded add-in list from the live app.
        $openAddins = @(Get-ExcelAddins -ExcelApp $ExcelApp)

        # This instance was just obtained via COM, so it is COM-reachable.
        $logs = if ($Log) { @($Log) } else { @() }
        Update-ExcelInstance -InstanceId $instancePid -Owner $Owner -Visible $isVisible -Reason $Reason -Intent $Intent -Workbooks $openWorkbooks -Logs $logs -Addins $openAddins -ComReachable $true
        return $instancePid
    } catch {
        Write-Warning "Failed to register Excel instance: $($_.Exception.Message)"
        return $null
    }
}

<#
.SYNOPSIS
Refresh the recorded workbooks/addins (and visibility) of a tracked instance
from its live COM Application object, preserving owner/reason/createdAt. Called
after opening workbooks/addins (and again just before teardown) so the registry
reflects what the instance actually holds — which lets the end-of-suite
assessment attribute a lingering ("zombie") instance to the workbooks (and
therefore the e2e test) that were open on it.
#>
function Refresh-ExcelInstance {
    param(
        [Parameter(Mandatory)][object]$ExcelApp,
        [int]$ProcessId = 0,
        [string]$TargetPath = '',
        [string]$Log = ''
    )

    try {
        # Resolve the pid from the live app when the caller didn't supply one.
        if ($ProcessId -le 0) {
            try { $ProcessId = Get-ExcelProcessId -ExcelApp $ExcelApp } catch { }
        }
        if ($ProcessId -le 0) { return }

        $openWorkbooks = @(Get-ExcelWorkbooks -ExcelApp $ExcelApp)
        $openAddins = @(Get-ExcelAddins -ExcelApp $ExcelApp)

        # Record the file the macro was asked to operate on (build/export/extract
        # open that file *inside* the VBA macro, so it is never visible via
        # Get-ExcelWorkbooks). Merging it here gives the registry a stable path
        # to attribute a lingering instance back to the test that used it.
        if ($TargetPath) {
            $normalizedTarget = $TargetPath -replace '\\', '/'
            if ($openWorkbooks -notcontains $normalizedTarget) {
                $openWorkbooks += $normalizedTarget
            }
        }

        $isVisible = $false
        try { $isVisible = [bool]$ExcelApp.Visible } catch {}

        Get-ExcelInstancesLock | Out-Null
        try {
            $doc = Read-ExcelRegistryDocument
            $active = @($doc.active)
            $found = $false

            for ($i = 0; $i -lt $active.Count; $i++) {
                if ([int]$active[$i].pid -eq $ProcessId) {
                    $active[$i].workbooks = @($openWorkbooks)
                    $active[$i].addins = @($openAddins)
                    if ($Log) {
                        $activeLogs = @($active[$i].logs) + @($Log)
                        $active[$i].logs = @($activeLogs | Where-Object { $_ } | Select-Object -Unique)
                    } elseif ($null -eq $active[$i].logs) {
                        $active[$i].logs = @()
                    }
                    $active[$i].visible = [bool]$isVisible
                    $found = $true
                    break
                }
            }

            if ($found) {
                Write-ExcelRegistryDocument -Active $active -Inactive $doc.inactive -ForceEvolutionSnapshot:([bool]$Log)
            }
        } finally {
            Release-ExcelInstancesLock
        }
    } catch {
        # Best-effort; never break a run over a refresh failure.
    }
}

<#
.SYNOPSIS
Untrack an Excel.Application instance from the registry before it is torn down.
If $ConfirmProcessId is supplied, only the matching entry is removed.
#>
function Unregister-ExcelInstance {
    param(
        [object]$ExcelApp,
        [int]$ProcessId = 0
    )

    $pidToRemove = $ProcessId
    if ($pidToRemove -le 0 -and $null -ne $ExcelApp) {
        try {
            $pidToRemove = Get-ExcelProcessId -ExcelApp $ExcelApp
        } catch {
            # Can't resolve; nothing to do.
            return
        }
    }

    if ($pidToRemove -gt 0) {
        Remove-ExcelInstance -InstanceId $pidToRemove
    }
}

<#
.SYNOPSIS
Return all live EXCEL.EXE processes with their PIDs, main window handle, title,
and whether the main window is visible. Used by the status script to flag
instances that are running but not recorded in the registry (rogue).
#>
function Get-RunningExcelInstances {
    $procs = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
    $result = @()
    foreach ($p in $procs) {
        $result += [pscustomobject]@{
            pid              = [int]$p.Id
            mainWindowHandle = [int64]$p.MainWindowHandle
            mainWindowTitle  = [string]$p.MainWindowTitle
            visible          = [bool]($p.MainWindowHandle -ne 0)
            startedAt        = $p.StartTime.ToString('o')
        }
    }
    return $result
}

<#
.SYNOPSIS
Guard against spawning a new Excel instance once too many EXCEL.EXE processes
are already running (leaked/rogue instances piling up from aborted runs).
Throws (does not call `Fail`, to keep this module decoupled from run.ps1's
output helpers) when the live count is at or above the limit.

`-Background` restricts the count to invisible (automation) instances only, so
a user's normally-open, visible Excel session doesn't eat into the automation
budget. Limit defaults to 12, overridable via VBA_MAX_EXCEL_INSTANCES.
#>
function Assert-ExcelInstanceLimit {
    param([bool]$Background = $false)

    $maxInstances = 12
    if ($env:VBA_MAX_EXCEL_INSTANCES) {
        $maxInstances = [int]$env:VBA_MAX_EXCEL_INSTANCES
    }

    $instanceCount = 0
    try {
        $running = @(Get-RunningExcelInstances)
        $instanceCount = if ($Background) {
            @($running | Where-Object { -not $_.visible }).Count
        } else {
            $running.Count
        }
    } catch {
        # Best-effort safeguard; never block a legitimate run over a count failure.
        return
    }

    if ($instanceCount -ge $maxInstances) {
        throw "Refusing to open a new Excel instance - $instanceCount EXCEL.EXE process(es) already running (limit $maxInstances). Run scripts/ps/Close-AllInvisibleExcelInstances.ps1 to clean up leaked instances."
    }
}

<#
.SYNOPSIS
Path to the instance-lifecycle debug log (%TEMP%\Excel-Instances\instances.log,
overridable via VBA_INSTANCE_LOG).
#>
function Get-InstanceLogPath {
    $path = $env:VBA_INSTANCE_LOG
    if (-not $path) {
        $path = Join-Path $env:TEMP 'Excel-Instances\instances.log'
    }
    return $path
}

<#
.SYNOPSIS
Append an instance-lifecycle log line (create/quit), only when
VBA_DEBUG_INSTANCES is set. Writes to a FILE, never stderr/stdout: run.ps1's
`toResult` on the Node side treats any stderr output as `success=false`, so
debug noise on stderr would turn a successful run into a false failure.
#>
function Write-InstanceLog {
    param([string]$Message)

    if (-not ($env:VBA_DEBUG_INSTANCES -match '^(1|true|yes)$')) {
        return
    }

    $count = -1
    try { $count = @(Get-RunningExcelInstances).Count } catch {}

    $logPath = Get-InstanceLogPath
    $stamp = (Get-Date).ToString('o')
    $line = "[vbapm-instances] $stamp pid=$PID count=$count $Message"

    try {
        Add-Content -LiteralPath $logPath -Value $line -ErrorAction Stop
    } catch {
        # Never let logging itself break a run.
    }
}

<#
.SYNOPSIS
Reconcile the registry against the live EXCEL.EXE process list. Already-tracked
entries keep their metadata; a live process not in the registry is added as a
minimal `rogue` entry only when `-TrackRogue` is supplied. Entries whose PID is
no longer live are moved to `inactive` (rather than dropped) so a later
assessment can detect a lingering process. Writes the stable snapshot shape
{ count, registered, invisible, active, inactive }.

This is the "stable layer": tracked entries stay current, while newly
discovered process-only rogues remain untracked unless explicitly persisted.
#>
function Sync-ExcelInstanceSnapshot {
    param([switch]$TrackRogue)

    $doc = Read-ExcelRegistryDocument
    $registered = @($doc.active)
    $inactive = @($doc.inactive)

    # Index existing entries by pid so we can enrich without duplicating.
    $byPid = @{}
    foreach ($e in $registered) {
        $byPid[[int]$e.pid] = $e
    }

    $live = @(Get-RunningExcelInstances)
    $livePids = @{}
    $snapshot = @()

    foreach ($p in $live) {
        $instancePid = [int]$p.pid
        $livePids[$instancePid] = $true

        if ($byPid.ContainsKey($instancePid)) {
            $entry = $byPid[$instancePid]
            # Refresh cheap, process-derived fields (visibility/title) even for
            # already-tracked entries so the snapshot stays current.
            $entry.visible = [bool]$p.visible
            if (-not $entry.windowTitle) { $entry.windowTitle = $p.mainWindowTitle }
            $snapshot += $entry
        } elseif ($TrackRogue) {
            # Live but untracked: record a minimal rogue entry so it shows up in
            # the registry when explicitly requested. Without -TrackRogue the
            # process remains visible in status output but is not persisted.
            $snapshot += [pscustomobject]@{
                id           = (New-ExcelInstanceId)
                number       = $null
                pid          = $instancePid
                owner        = 'rogue'
                visible      = [bool]$p.visible
                reason       = 'rogue'
                comReachable = $false
                windowTitle  = if ([string]::IsNullOrEmpty($p.mainWindowTitle)) { $null } else { $p.mainWindowTitle }
                createdAt    = $p.startedAt
                workbooks    = $null
                addins       = $null
                startedAt    = $p.startedAt
            }
        }
    }

    # Entries in `active` whose PID is no longer live are deactivated (moved to
    # `inactive`) rather than dropped, so a later assessment can tell whether
    # the process actually exited or is lingering.
    $dead = @($registered | Where-Object { -not $livePids.ContainsKey([int]$_.pid) })
    foreach ($entry in $dead) {
        $inactive += (ConvertTo-InactiveEntry -Entry $entry -DeactivateReason 'dead-process')
    }

    # Keep the inactive list bounded to the most recent entries.
    $maxInactive = 100
    if ($inactive.Count -gt $maxInactive) {
        $inactive = @($inactive | Select-Object -Last $maxInactive)
    }

    Write-ExcelRegistryDocument -Active $snapshot -Inactive $inactive
    return $snapshot
}

<#
.SYNOPSIS
Enumerate every COM-reachable Excel.Application instance via the Running Object
Table (ROT), de-duplicated by window handle. Bind every ROT moniker before
inspecting it. Excel workbook monikers can be bare names such as `Book1`, so
filtering display names before binding misses valid Excel applications.

Excel can register either an Application moniker or a Workbook moniker, so
workbook objects are resolved through their parent Application before
deduplication. A user-launched instance is often NOT present in the ROT (hence
GetActiveObject limitations) — this only returns instances Excel chose to
register.
#>
function Get-ExcelApplications {
    $rotClass = 'RotHelperExcel'
    if (-not ($rotClass -as [type])) {
        try {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class RotHelperExcel {
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable rot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(uint reserved, out IBindCtx ctx);

    public static System.Collections.Generic.List<object> GetExcelInstances() {
        var list = new System.Collections.Generic.List<object>();
        IRunningObjectTable rot = null;
        IBindCtx ctx = null;
        IEnumMoniker monikers = null;
        if (GetRunningObjectTable(0, out rot) != 0 || rot == null) return list;
        if (CreateBindCtx(0, out ctx) != 0 || ctx == null) {
            Marshal.ReleaseComObject(rot);
            return list;
        }
        IMoniker[] fetched = new IMoniker[1];
        try {
            rot.EnumRunning(out monikers);
            if (monikers == null) return list;
            monikers.Reset();
            while (monikers.Next(1, fetched, IntPtr.Zero) == 0) {
                var moniker = fetched[0];
                object obj = null;
                try {
                    try {
                        if (rot.GetObject(moniker, out obj) == 0 && obj != null) {
                            list.Add(obj);
                            obj = null;
                        }
                    } catch {}
                } finally {
                    if (obj != null) { try { Marshal.ReleaseComObject(obj); } catch {} }
                    if (moniker != null) { try { Marshal.ReleaseComObject(moniker); } catch {} }
                }
            }
        } finally {
            if (monikers != null) Marshal.ReleaseComObject(monikers);
            if (ctx != null) Marshal.ReleaseComObject(ctx);
            if (rot != null) Marshal.ReleaseComObject(rot);
        }
        return list;
    }
}
"@
        } catch {
            # ROT helper failed to load; return nothing.
            return @()
        }
    }

    $apps = @()
    try {
        $apps = @([RotHelperExcel]::GetExcelInstances())
    } catch {
        return @()
    }

    # Resolve workbook monikers to their parent Application, then de-duplicate
    # by Hwnd because the same Application can appear under multiple monikers.
    $seen = @{}
    $deduped = @()
    foreach ($candidate in $apps) {
        if ($null -eq $candidate) { continue }

        $app = $null
        $isApplication = $false
        try {
            if ([string]$candidate.Name -eq 'Microsoft Excel') {
                $app = $candidate
                $isApplication = $true
            }
        } catch {}

        if ($null -eq $app) {
            try { $app = $candidate.Application } catch { $app = $null }
        }

        if ($null -eq $app) {
            try {
                $null = $candidate.Workbooks.Count
                $app = $candidate
                $isApplication = $true
            } catch {}
        }

        if ($null -eq $app) {
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($candidate) | Out-Null } catch {}
            continue
        }

        try {
            if ([string]$app.Name -ne 'Microsoft Excel') {
                if (-not $isApplication) {
                    try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null } catch {}
                }
                try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($candidate) | Out-Null } catch {}
                continue
            }
        } catch {
            if (-not $isApplication) {
                try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null } catch {}
            }
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($candidate) | Out-Null } catch {}
            continue
        }

        $key = $null
        try { $key = [string]$app.Hwnd } catch { $key = [string]$app.GetHashCode() }
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $deduped += $app
        }

        if (-not $isApplication) {
            try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($candidate) | Out-Null } catch {}
        }
    }
    return @($deduped)
}

<#
.SYNOPSIS
Search all COM-reachable Excel instances for a workbook that is already open at
the given path. Returns $null if not found, otherwise a pscustomobject with
`.App` (the Excel.Application) and `.Workbook` (the open workbook).

Matches by normalized, case-insensitive full path so the same file opened under
a different casing or slash style is still recognized.
#>
function Find-OpenWorkbook {
    param([Parameter(Mandatory)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $normalized = $fullPath -replace '\\', '/'
    $normalized = $normalized.ToLowerInvariant()

    $apps = @(Get-ExcelApplications)
    foreach ($app in $apps) {
        if ($null -eq $app) { continue }
        try {
            foreach ($wb in $app.Workbooks) {
                try {
                    $wbFull = ([string]$wb.FullName) -replace '\\', '/'
                    if ($wbFull.ToLowerInvariant() -eq $normalized) {
                        return [pscustomobject]@{ App = $app; Workbook = $wb }
                    }
                } catch {
                    # skip this workbook
                }
            }
        } catch {
            # skip this instance
        }
    }

    return $null
}
