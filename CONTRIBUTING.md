# Contributing

Thanks for contributing to vbapm.

## Development Setup

### Branch Strategy

The default branch is `main`. Make sure your feature branches branch out from it.

### Initialize Submodules

The repository uses Git submodules. After cloning, ensure they are synced and initialized:

```powershell
git submodule update --init --recursive
```

This will pull the following submodules:

| Submodule | Path |
|-----------|------|
| `installer` | `installer/` |
| `VBA-on-GitHub` | `src/actions/templates/VBA-on-GitHub/` |

### Install and Build

1. Run `pnpm install`
2. Run `pnpm run lefthook:install` — sets up pre-commit hooks for formatting and linting.
3. Run `pnpm run format`
4. Run `pnpm run build:cli` — builds the CLI/library into `lib/` and ensures the vendored Node runtime is available.
4. Run `pnpm run build:addins` — creates `addins\build\vbapm.xlam` which performs workbook/VBA operations from inside Office.

### Development Checks

Run all checks in one command:

```powershell
pnpm run dev
```

This runs type checking, formatting checks, unit tests and a CLI build. Integration tests (e2e) are not included — see the next section.

The `dev` command is composed of:

- `pnpm run deps` — install dependencies
- `pnpm run typecheck` — TypeScript type checking
- `pnpm run format:check` — check code formatting
- `pnpm run test` — run unit tests
- `pnpm run build:cli` — build the CLI

> **Note:** `build:addins` is not included in `dev`. Run it separately before e2e tests or whenever the VBA addin source changes.

### Generated artifact freshness

`pnpm run build` and `pnpm run build:check` refresh the compiled `lib/` output
and run the add-in freshness check. The add-in check watches only VBA and XML
files under `addins/`, so TypeScript changes do not start Excel by default.
VBA or XML changes make the add-in stale and trigger its rebuild.

When a TypeScript change is known to affect the add-in build or the bootstrap
workbook, force the full freshness check:

```powershell
pnpm run build:check -- --force
```

The `--force` pass forces the add-in rebuild and checks the bootstrap artifact
after refreshing `lib/`. Run `pnpm run build:addins` directly when you need to
build the add-in independently. The bootstrap refresh workflow uses `--force`
on CI.

### Test Commands

- **Unit tests:** `pnpm run test`
- **E2E tests (visible Excel):** `pnpm run test:e2e`
- **E2E tests (hidden Excel):** `pnpm run test:e2e:background`

To run a specific test, use the `--testNamePattern` flag:

```powershell
pnpm run test -- --testNamePattern="build standard project"
```

## Committing

Before committing, ensure all tests pass and formatting is correct (`pnpm run format:check`). Write commit messages in the [conventional commit](https://www.conventionalcommits.org/) format.

## Integration Tests (e2e)

The e2e tests exercise the full pipeline — from the CLI binary through the Excel addin — against real Excel instances. They validate end-to-end scenarios like project export, import, build and dependency resolution.

### Why `:background`?

The e2e tests interact with Excel via the COM automation API (through the vbapm.xlam addin). Without the `:background` option, Excel opens visibly — you'll see workbook windows flash open and close as each test runs. This is expected behavior but can interfere with any work you have open in Excel.

The `:background` variant (`pnpm run test:e2e:background`) passes an explicit background option to the CLI, which creates a new hidden Excel instance via COM (`Excel.Visible = false`) instead of attaching to an already-running visible instance. This prevents the window flashing and keeps your existing Excel work undisturbed.

### Performance

A full e2e test run takes roughly **60-90 seconds** to complete. If you're using the background variant, don't worry if nothing appears in the terminal for a while — that's normal. The tests are running; they just don't pollute your terminal with intermediate output.

To inspect temporary test folders after a run, set `KEEP_E2E_TMP=1` before running the tests.

## Local CLI Installation with `devinstall.ps1`

To test `vba` / `vbapm` as an installed command (rather than via `bin/vba`), use the local dev installer. It copies the current build artifacts into `%APPDATA%\vbapm` and adds the `bin` directory to your user `PATH`.

### Prerequisites

Build the CLI and optionally the add-in before installing:

```powershell
pnpm run build:cli
pnpm run build:addins   # optional — needed only if you use add-in commands
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
pnpm run build:cli
.\installer\devinstall.ps1
```

### Uninstall

Delete the installed directory and remove the `%APPDATA%\vbapm\bin` entry from your user `PATH`:

```powershell
Remove-Item "$env:APPDATA\vbapm" -Recurse -Force
```

> **Note:** The `devinstall.ps1` script requires an unrestricted or bypassed execution policy. Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` if you see a PowerShell execution policy error.

## Verbose e2e Logs

e2e command output can be echoed even when tests succeed.

Use one of these options:

- Pass Jest verbose through pnpm args:
  - PowerShell: `pnpm run test:e2e:background -- --verbose`
- Set the environment variable used by the e2e helper:
  - PowerShell: `$env:E2E_VERBOSE=1; pnpm run test:e2e:background`

This prints each invoked e2e command plus its stdout and stderr, which helps compare local runs with CI logs.

## Working with Git Worktrees

When working with multiple branches in parallel, you can create a worktree under `worktrees/`. The following setup steps are required before running tests in a worktree.

### 1. Install `node_modules`

Worktrees share the git repository but each needs its own `node_modules`. Run from inside the worktree:

```powershell
cd worktrees/<branch-name>
pnpm install
```

### 2. Build the CLI (`lib/`)

The e2e tests invoke the local `bin/vba` binary which requires `lib/` to be built:

```powershell
pnpm run build:cli
```

### 3. Build the addin (`addins/build/vbapm.xlam`)

The e2e tests rely on the Excel addin to import/export VBA code. It must be built once per worktree:

```powershell
pnpm run build:addins
```

### 4. Running tests from a worktree

The root `jest.config.mjs` and `e2e.config.mjs` use `testPathIgnorePatterns: ["<rootDir>/worktrees/"]`. Because `<rootDir>` resolves to the worktree root when Jest is invoked from inside it, the pattern matches nothing — so tests are correctly discovered.

Always `cd` into the worktree before running `pnpm test` or `pnpm run test:e2e:*`. Use `Push-Location` for background terminal commands so the current working directory is preserved:

```powershell
Push-Location worktrees/<branch-name>; pnpm run test:e2e:updateSnapshots 2>&1 | Tee-Object "$env:TEMP\e2e.log"; Pop-Location
```

### 5. Snapshot files

Snapshot files live under `tests/__snapshots__/` which is tracked by git inside the worktree. Each worktree has its own independent snapshot state — run `pnpm run test:e2e:updateSnapshots` from the worktree to baseline them after making changes.
