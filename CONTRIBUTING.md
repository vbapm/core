# Contributing

Thanks for contributing to vbapm.

## Development Setup

1. Run `npm install`
2. Run `npm run format`
3. Run `npm run build:cli`
4. Run `npm run build:addins`

## Local CLI Installation with `devinstall.ps1`

To test `vba` / `vbapm` as an installed command (rather than via `bin/vba`), use the local dev installer. It copies the current build artifacts into `%APPDATA%\vbapm` and adds the `bin` directory to your user `PATH`.

### Prerequisites

Build the CLI and optionally the add-in before installing:

```powershell
npm run build:cli
npm run build:addins   # optional — needed only if you use add-in commands
```

### Install

From the repository root, run:

```powershell
.\installer\devinstall.ps1
```

The script copies these directories to `%APPDATA%\vbapm`:

- `bin` — entry-point scripts (`vba`, `vbapm`, `vba.cmd`, etc.)
- `lib` — compiled CLI code
- `vendor` — bundled Node.js runtime
- `run-scripts` — AppleScript / PowerShell run helpers
- `addins` — Excel add-in (if built)

After installation the `vba` command is available in new terminals:

```powershell
vba --help
```

### Reinstall

Rebuild and re-run the installer to pick up changes:

```powershell
npm run build:cli
.\installer\devinstall.ps1
```

### Uninstall

Delete the installed directory and remove the `%APPDATA%\vbapm\bin` entry from your user `PATH`:

```powershell
Remove-Item "$env:APPDATA\vbapm" -Recurse -Force
```

> **Note:** The `devinstall.ps1` script requires an unrestricted or bypassed execution policy. Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` if you see a PowerShell execution policy error.

## Testing

1. Run unit tests with `npm test`
2. Run end-to-end tests with `npm run test:e2e` or `npm run test:e2e:background`
3. To keep temporary e2e folders for inspection, set `KEEP_E2E_TMP=1` before running tests

## Verbose e2e Logs

e2e command output can be echoed even when tests succeed.

Use one of these options:

- Pass Jest verbose through npm args:
  - PowerShell: `npm run test:e2e:background -- --verbose`
- Set the environment variable used by the e2e helper:
  - PowerShell: `$env:E2E_VERBOSE=1; npm run test:e2e:background`

This prints each invoked e2e command plus its stdout and stderr, which helps compare local runs with CI logs.
