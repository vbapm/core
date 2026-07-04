# Open/Close via VBA Host Macros — Design Decision

## Motivation

The `close` command needs to check whether a workbook has unsaved changes before closing — but the logic for detecting unsaved changes varies across VBA-enabled applications (Excel, Word, PowerPoint, etc.). Implementing this check in separate PowerShell and AppleScript bridge scripts for each host application is probaly more complex than needed. VBA macros in the vbapm addin provide a single, unified implementation that is shared across platforms and naturally handles application-specific differences.

## Chosen Architecture

| Operation | Mechanism | Rationale |
|-----------|-----------|-----------|
| **Open** | Node.js `open` package → OS default handler | Instant, addin-independent, works for all file types |
| **Close** | `run()` → bridge → addin → VBA macro (`Build.CloseFile`) | Single VBA impl for all hosts; uniform error handling |
| **Check Saved** | `run()` → bridge → addin → VBA macro (`Build.CheckFileSaved`) | Same pipeline as close; host-specific `.Saved` logic in VBA |

### Flows

**OPEN** (unchanged — OS-native):
```
CLI: vbapm open [--target=TYPE]
  ↓
getTargetPath() → resolves build/target.xlsx
  ↓
openTarget(path) → Node.js open(path, { wait: true })
  ↓
System default app opens the file
```

**CLOSE** (via VBA addin):
```
CLI: vbapm close [--target=TYPE] [--save] [--force]
  ↓
resolveTargetPath()
  ↓
If not --force and not --save:
  run("excel", addin, "Build.CheckFileSaved", {file})
  └─ Addin → Application.Workbooks(file).Saved
  └─ If unsaved → throw CliError (suggest --save or --force)
  ↓
run("excel", addin, "Build.CloseFile", {file, save})
  └─ Addin → Application.Workbooks(filename).Close(save)
  └─ Returns JSON result
```

## VBA Macros

Both macros are in `addins/src/Build.bas` and follow the same pattern as existing addin macros.

### Build.CloseFile

Takes a JSON string `{ file, save }`. Looks up the workbook by filename in `Application.Workbooks()`. If found, closes it with or without saving. If not found, reports "File is not open" (no error).

### Build.CheckFileSaved

Takes a JSON string `{ file }`. Returns `"saved:true"` or `"saved:false"` via the standard Output messages. If the workbook is not open, treats it as saved (consistent with the behaviour that an unopened file cannot have unsaved changes).

## Legacy Bridge Script Code

The `Close` and `CheckSaved` functions were removed from `run.ps1` and `run.applescript` in commit `e1f486b`. Both scripts now only handle macro execution — the same code path used by `build`, `test`, and `run` commands.

## Trade-offs

| Factor | Assessment |
|--------|------------|
| **Infrastructure uniformity** | ✅ Close/CheckSaved now use the same `run()` pipeline as build/test/run |
| **Cross-platform maintenance** | ✅ VBA code is shared; only the runner scripts differ by platform |
| **Error handling** | ✅ VBA `Err` descriptions surface as CLI error messages, matching other commands |
| **Addin dependency** | ⚠️ Close requires the addin to be installed; open does not |
| **Performance** | ⚠️ Close is slightly slower (addin load overhead) vs direct COM, but negligible in practice |
| **Host extensibility** | ✅ Adding support for Word/PowerPoint close only requires VBA macros, not new bridge script code |

Related Pull Request: https://github.com/vbapm/core/pull/63

## Historical Note

The `close` command was initially implemented with direct COM/AppleScript calls in the bridge scripts (`run.ps1` / `run.applescript`), bypassing the addin entirely. This approach was fast and addin-independent — it attached to any running Excel instance and called `Workbook.Close()` directly.

- `a50369b` — Consolidated close logic into the bridge scripts with `Close` and `CheckSaved` functions and `-Close` / `-CheckSaved` PowerShell switches.
- `e1f486b` — Removed the bridge-script close logic and migrated to VBA macros (`Build.CloseFile` / `Build.CheckFileSaved`), routing through the addin's `run()` pipeline instead.

The motivation for the migration was that detecting unsaved changes in a host-agnostic way (Excel, Word, PowerPoint, etc.) requires VBA code specific to each application, and centralizing that logic in the addin is more maintainable than duplicating it across PowerShell and AppleScript bridge scripts.
