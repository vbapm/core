# PowerShell Scripts Reference

This document lists the PowerShell scripts in [`scripts/ps/`](../../scripts/ps/), with a
brief description of each.

## Excel instance coordination

These scripts implement cross-agent coordination for Excel instances. When multiple
automation agents (or a user) work in the same repository, they all need to drive Excel
via COM — which can collide if instances are not tracked. This suite maintains a central
registry under `%TEMP%\Excel-Instances\`:

- `instances.json` — the registry (a JSON array of tracked instance entries).
- `instances.lock` — a file lock used to serialize edits to `instances.json`.

Each registry entry describes one Excel `EXCEL.EXE` process and includes:

| Field | Description |
|---|---|
| `id` | Stable lowercase-alphanumeric hash (e.g. `a1b2c3d4`), used to close that instance. |
| `pid` | Windows process id of the `EXCEL.EXE` process. |
| `owner` | Who created it, e.g. `terminal-12345`. |
| `createdAt` | ISO-8601 creation timestamp. |
| `visible` | Whether the Excel window is visible. |
| `reason` | Why it was created, e.g. `e2e`, `vbapm-run`, `unknown`. |
| `workbooks` | Array of open workbook paths (forward-slash normalized). |

> **Note:** Excel registers its `Application` COM object in the Running Object Table
> (ROT) under CLSID monikers (not the friendly `"Excel.Application"` name), and a
> user-launched instance is often **not** registered in the ROT at all. See
> [COM instance discovery](#com-instance-discovery) below.

The scripts below dot-source [`Excel-InstanceRegistry.ps1`](#excel-instanceregistryps1),
which contains the shared functions. Except for the two standalone diagnostic/utility
scripts, all of these are **Windows-only**.

| Script | Description |
|---|---|
| [`Excel-InstanceRegistry.ps1`](#excel-instanceregistryps1) | **Library.** Dot-sourced by the others. Registry read/write (lock-guarded), lock acquire/release, PID/visibility/workbook capture, id-hash generation, and close-by-id. |
| [`Get-ExcelInstancesStatus.ps1`](#get-excelinstancesstatusps1) | Report the registry + live `EXCEL.EXE` processes. Flags rogue instances (running but untracked) and orphaned entries (tracked but dead). `-Json` for machine-readable output, `-FailOnRogue` to exit non-zero on rogues. |
| [`Sync-ExcelInstances.ps1`](#sync-excelinstancesps1) | Discover all running Excel instances and (re)write the registry with current visibility + workbooks. `-PruneDead` also drops stale entries. |
| [`Register-ExcelInstance.ps1`](#register-excelinstanceps1) | Manually track an instance by `-ProcessId`. |
| [`Unregister-ExcelInstance.ps1`](#unregister-excelinstanceps1) | Untrack an instance by `-ProcessId`. |
| [`Close-ExcelInstance.ps1`](#close-excelinstanceps1) | Close (quit) an instance by its hash `-Id` or `-ProcessId`, then remove it from the registry. |
| [`Acquire-ExcelInstancesLock.ps1`](#acquire-excelinstanceslockps1) | Acquire the registry lock (held until this process exits or `-Release`). |
| [`Release-ExcelInstancesLock.ps1`](#release-excelinstanceslockps1) | Explicitly release the registry lock. |
| [`Ensure-ExcelInstancesClean.ps1`](#ensure-excelinstancescleanps1) | Guard run before the e2e suite. Fails fast if a rogue instance is running. |

## Other scripts

| Script | Description |
|---|---|
| [`Diagnose-VBIDETypeLib.ps1`](#diagnose-vbidetypelibps1) | Check whether the VBIDE type library (`VBE6EXT.OLB` / `VBE7.DLL`) is registered, and report its GUID/version/paths. |
| [`Trim-VbaTrailingWhitespace.ps1`](#trim-vbatrailingwhitespaceps1) | Strip trailing spaces/tabs from `.bas`/`.cls`/`.frm`/`.vba`/`.doccls` files while preserving the original encoding. |

---

## `Excel-InstanceRegistry.ps1`

Dot-sourced **library** — not meant to be run directly. Provides the shared building
blocks used by all the other coordination scripts:

- `Get/Read/Write-ExcelInstances` — registry path + JSON read/write.
- `Get/Release-ExcelInstancesLock` — atomic lock file (stale-lock stealing).
- `Register/Unregister-ExcelInstance` — track/untrack an instance.
- `Update/Remove-ExcelInstance` — low-level upsert/delete by PID.
- `Find/Remove/Close-ExcelInstanceById` — lookup/remove/close by hash id.
- `Get-ExcelProcessId`, `Get-ExcelWorkbooks`, `New-ExcelInstanceId` — intrinsics.

## `Get-ExcelInstancesStatus.ps1`

```powershell
.\scripts\ps\Get-ExcelInstancesStatus.ps1            # human-readable report
.\scripts\ps\Get-ExcelInstancesStatus.ps1 -Json      # machine-readable
.\scripts\ps\Get-ExcelInstancesStatus.ps1 -FailOnRogue   # exit 1 if rogues exist
```

## `Sync-ExcelInstances.ps1`

```powershell
.\scripts\ps\Sync-ExcelInstances.ps1                 # refresh registry from live instances
.\scripts\ps\Sync-ExcelInstances.ps1 -PruneDead      # also drop stale entries
```

## `Register-ExcelInstance.ps1`

```powershell
.\scripts\ps\Register-ExcelInstance.ps1 -ProcessId 1234 -Reason e2e
.\scripts\ps\Register-ExcelInstance.ps1 -ProcessId 1234 -Reason manual -Visible
```

## `Unregister-ExcelInstance.ps1`

```powershell
.\scripts\ps\Unregister-ExcelInstance.ps1 -ProcessId 1234
```

## `Close-ExcelInstance.ps1`

```powershell
.\scripts\ps\Close-ExcelInstance.ps1 -Id a1b2c3d4    # close by hash (preferred)
.\scripts\ps\Close-ExcelInstance.ps1 -ProcessId 1234 # close by pid
```

## `Acquire-ExcelInstancesLock.ps1`

```powershell
.\scripts\ps\Acquire-ExcelInstancesLock.ps1          # acquire
.\scripts\ps\Acquire-ExcelInstancesLock.ps1 -Release # release
```

## `Release-ExcelInstancesLock.ps1`

```powershell
.\scripts\ps\Release-ExcelInstancesLock.ps1
```

## `Ensure-ExcelInstancesClean.ps1`

Run immediately before the e2e integration suite. Inspected by
`scripts/ensure-excel-instances-clean.ts`, which is pre-pended to all `test:e2e*` npm
scripts and is a **no-op on non-Windows** platforms.

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\ps\Ensure-ExcelInstancesClean.ps1
```

## `Diagnose-VBIDETypeLib.ps1`

```powershell
.\scripts\ps\Diagnose-VBIDETypeLib.ps1
```

## `Trim-VbaTrailingWhitespace.ps1`

```powershell
.\scripts\ps\Trim-VbaTrailingWhitespace.ps1 -Path .\src
```

---

## COM instance discovery

There is a well-known COM limitation that affects everything in this suite:

- `Marshal.GetActiveObject("Excel.Application")` (or `GetActiveObject`) returns **at most
  one** instance — whichever Excel registered itself in the ROT first.
- Excel registers its `Application` object in the ROT under CLSID monikers like
  `!{00024500-0000-0000-C000-000000000046}`, **not** under a friendly name.
- A user-launched instance (e.g. double-clicking a `.xlsm`) is frequently **not** in the
  ROT at all, so it cannot be reached via `GetActiveObject`.
- `AccessibleObjectFromWindow` (OBJID_NATIVEOM) returns `E_FAIL` for Excel windows
  (unlike Word/Outlook).

Consequences and how the scripts handle them:

1. `run-scripts/run.ps1` only ever attaches via `GetActiveObject` for the non-background
   path, and otherwise creates its own instance.
2. `Sync-ExcelInstances.ps1` enumerates the ROT **by CLSID** for COM-reachable instances,
   then falls back to **process + window-title** enumeration for instances that aren't
   ROT-registered (e.g. a visible user session). For the latter, the workbook list is
   inferred from the main-window title (best-effort).
3. Ownership (`terminal-<pid>`) is stamped **at creation time**, because COM launches
   Excel via `svchost.exe` (the COM SCM), which breaks the parent-process chain — so the
   creator cannot be reliably inferred afterwards.
