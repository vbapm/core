# `vbapm update` Command

Import VBA source directly into a built target — including one that is currently open in Excel — without rebuilding the file from XML.

## Usage

```
vbapm update [options]

Options:
  --open          Open the target in Excel after updating (if not already open)
```

## How It Works

```
loadProject()
  └─ getTarget()
  └─ fetchDependencies()
        │
        ▼
loadFromProject()          ← reads src/*.bas / .cls / .frm
stageBuildGraph()          ← copies files to a temp staging dir
        │
        ▼
importTarget(target, info, builtFilePath, options)
  └─ Build.ImportGraph VBA macro (via run.ps1)
       ├─ if file already open in Excel → attaches to it
       ├─ if file not open             → opens build/file from disk
       ├─ imports all src components
       ├─ updates references
       └─ Document.Save()
```

The key difference from `build`: `update` operates on `build/<target>` directly instead of a staged copy, and performs **no** zip/unzip, backup, or move operations.

## Files to Add / Modify

| File | Change |
|---|---|
| `src/bin/vbapm-update.ts` | **New** — CLI entry point |
| `src/actions/update-project.ts` | **New** — `updateProject()` action |
| `src/bin/vbapm.ts` | **Modify** — register `update` in the `commands` map |
| `run-scripts/run.ps1` | **Modify** — accept `keepOpen` parameter in `Dispose()` |

## Implementation Details

### `src/actions/update-project.ts`

Reuses `importTarget` (already exported from `src/targets/build-target.ts`) — that function does exactly: load build graph → stage → call `Build.ImportGraph`. The only difference is that we pass the live built file path directly instead of a staged copy.

```typescript
export async function updateProject(options: UpdateOptions = {}): Promise<string> {
  // [1/2] Load project
  const project = await loadProject();
  const { target } = getTarget(project, options.target);
  const dependencies = await fetchDependencies(project);

  // Guard: built file must exist (update does not create it)
  const builtFile = join(project.paths.build, target.filename);
  if (!await pathExists(builtFile)) {
    throw new CliError(ErrorCode.UpdateTargetNotBuilt, `...`);
  }

  // [2/2] Import VBA into target (live or on-disk)
  await importTarget(target, { project, dependencies }, builtFile, options);

  return builtFile;
}
```

### `run-scripts/run.ps1` — `--open` behaviour in `Dispose()`

The `--open` flag is passed as an extra argument to `run.ps1` (or forwarded via an existing args slot) so that `Dispose()` can make the right call. The rules are:

| State when update starts | `--open` not set | `--open` set |
|---|---|---|
| File already open in Excel | Leave file open | Leave file open |
| Excel open, file not open | Close workbook; leave Excel running | Leave file open; leave Excel running |
| Excel not running | Close workbook; quit Excel | Leave file open; leave Excel running |

In short: a file that was **already open** is never touched by `Dispose()`. Everything else follows `--open`.

The `Dispose()` logic becomes:

```powershell
[void] Dispose([bool]$KeepOpen) {
    # A file that was open before we started is never closed
    $closeWorkbook = -not $this.WorkbookWasOpen -and -not $KeepOpen

    if ($closeWorkbook -and $null -ne $this.Workbook) {
        $this.Workbook.Close($true)
        $this.Workbook = $null
    }
    # Quit Excel only if we launched it AND we're not keeping the file open
    if (-not $this.ExcelWasOpen -and -not $KeepOpen -and $null -ne $this.App) {
        $this.App.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.App) | Out-Null
        $this.App = $null
    }
}
```

`update-project.ts` passes `options.open` down through `importTarget` so the value reaches the `run.ps1` invocation.

### Error codes to add to `src/errors.ts`

- `SyncTargetNotBuilt` — no built file found; user must run `build` first

### Help text registration in `src/bin/vbapm.ts`

```typescript
update: async () => (await import("./vbapm-update")).default,
```

And add `update` to the `Commands:` section of the help string.

