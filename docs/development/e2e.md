Here's a summary of the e2e testing process in this repo.

## What e2e CLI tests exercise

The end-to-end tests (excel.e2e.ts + helper in execute.ts) run the full pipeline — CLI binary → Excel COM automation via the `vbapm.xlam` add-in — against real Excel on Windows. They cover build, `export`, `extract`, `update`, `new`, `close`, `run`, and worksheet/drawing edge cases, comparing output against Jest snapshots.

## How the tests are invoked

From package.json:

| Script | What it does |
|---|---|
| `pnpm run test:e2e` | `VBA_BACKGROUND_BUILD=0` (visible Excel), **spawns the real CLI** per command |
| `pnpm run test:e2e:in-process` | visible Excel + `E2E_IN_PROCESS=1`, **dispatches CLI commands in-process** (see below) |
| `pnpm run test:e2e:background` | `VBA_BACKGROUND_BUILD=1` (hidden Excel) — what we just ran |
| `pnpm run test:e2e:updateSnapshots` | same as background + `--updateSnapshot` |
| `pnpm run test:e2e:multilang` | separate multilang config |

Each script is prefixed with `pnpm run build:check` → `node scripts/ensure-fresh-build.ts`, which rebuilds lib if sources are newer than the build output (that's the `[ensure-fresh-build] lib/ is stale … rebuilding` line you saw).

The actual test runner is `jest --config e2e.config.mjs --runInBand` (serial, so no parallel Excel collisions).

### Two ways to drive the CLI: spawned vs in-process

The `execute()` helper can drive each command one of two ways, selected by the
`E2E_IN_PROCESS` env var (the script sets it; the helper reads it):

1. **Spawned CLI** (`E2E_IN_PROCESS` unset — `test:e2e`, `test:e2e:background`).
   Each command is run as a **fresh `node lib/vbapm.js` process** via
   `child_process.exec`, exactly like a real user invocation. This is the
   default and the most faithful reproduction of the CLI.

2. **In-process dispatch** (`E2E_IN_PROCESS=1` — `test:e2e:in-process`).
   Commands are dispatched **inside the Jest process** by importing the CLI's
   bin command modules (`src/bin/vbapm-*.ts`) directly and invoking their
   default exports — no `cmd.exe`, no `vba.cmd`, no `node` child process per
   command. This skips process spawn overhead, so the suite runs faster
   (~125s vs ~145s). To keep `stdout`/`stderr` byte-identical to the spawned
   CLI, `executeInProcess` overrides `console.log`/`warn`/`error` (Jest
   otherwise intercepts them) and stubs the ESM-only `open` package
   (`tests/__helpers__/open-stub.ts`, mapped in `e2e.config.mjs`). The `open`
   command itself always falls back to the spawned CLI.

## Bounded Excel instance pool (parallel runs)

Every `run()` call (the `vba run` / macro-execution path) funnels through a
**process-local counting semaphore** (`src/utils/semaphore.ts` +
`src/utils/excel-pool.ts`) so that at most `VBA_EXCEL_POOL_SIZE` Excel instances
are alive at once:

- Default pool size is **4**; set `VBA_EXCEL_POOL_SIZE` to a different value to
  tune it (e.g. `1` for fully serialized Excel access).
- A caller that finds all slots busy **waits** (async) until one frees, rather
  than spawning yet another Excel instance. Reuse of an *already-open workbook*
  is handled separately by `Find-OpenWorkbook` in the PowerShell bridge.
- The slot is always released (in `finally`) even if the macro run throws.

This is what makes parallel execution viable: instead of `--runInBand` (one test
at a time), Jest workers can run concurrently and still respect a bounded Excel
footprint. Note the current npm scripts still pass `--runInBand` and set no pool
size; running in parallel is opt-in and not yet wired into a script.

## Two distinct execution paths *inside* the tests

This is the crucial detail:

1. **`execute(cwd, "build")`** — spawns the compiled vba binary as a **fresh Node process** per command. Used for build, `export`, `extract`, `new`, etc.

2. **`run("excel", file, "Validation.Validate")`** — calls the `run()` function **in-process** from the `"vbapm"` library loaded into the Jest process. Used to actually execute a macro inside a built workbook.

## Call chain and process boundaries

```text
Jest process (repo checkout)
├─ execute() helper in tests/__helpers__/execute.ts
│  └─ child_process.exec
│     └─ cmd.exe (Windows shell)
│        └─ vba.cmd shim
│           └─ node.exe
│              └─ lib/vbapm.js
│                 └─ powershell.exe
│                    └─ EXCEL.EXE
│                       └─ vbapm.xlam add-in
│                          └─ Build.ImportGraph / Build.CreateDocument / Build.ExportTo
└─ run() helper in tests/__helpers__/execute.ts
   └─ lib/vbapm.js (in-process; dispatches src/bin/vbapm-*.ts)
      └─ powershell.exe
         └─ EXCEL.EXE
            └─ vbapm.xlam add-in
               └─ Build.ImportGraph / Build.CreateDocument / Build.ExportTo
```

`child_process.exec` runs the command through the platform shell, which is why `execute()` follows the `cmd.exe` → `vba.cmd` → `node.exe` chain on Windows.

The CLI-side PowerShell bridge that actually launches Excel lives in [src/utils/run.ts](src/utils/run.ts).

**Key process boundaries:**

- **Jest (repo process)** — drives tests.
  - **Path A — `execute()`**: calls `child_process.exec`, which on Windows routes the command through **`cmd.exe`**. `cmd.exe` resolves the extensionless `vba` path to `vba.cmd` (via `PATHEXT`), which launches `node.exe` running `lib/vbapm.js`.
  - **Path B — `run()` helper**: calls the library **in-process** (no `cmd.exe`, no new `vba` Node process).
- **`cmd.exe`** — the Windows shell hosting the `vba.cmd` batch shim. Only in path A.
- **`vba.cmd` shim** — a batch wrapper that invokes `node.exe --no-warnings lib/vbapm.js` (or `vendor/node.exe` if bundled).
- **`node.exe`** — the actual runtime for the CLI; loads `lib/vbapm.js`.
- **`powershell.exe`** — the COM bridge (`run.ps1`), spawned by `lib/vbapm.js` (via `run()` in `src/utils/run.ts`) only when a macro must run. This is the only layer that talks to Excel.
- **`EXCEL.EXE`** — out-of-process COM server hosting the `vbapm.xlam` add-in; `Application.Run` executes the build/export/import macros.
- **Coordination registry** — a side-channel (not part of the macro relay) that tracks instances across all agents for cross-agent awareness.

> **Note:** `execute()` (via `child_process.exec`) spawns `cmd.exe`, which runs `vba.cmd`, which launches `node.exe` running `lib/vbapm.js`; that is where `powershell.exe` is finally spawned. The `run()` helper skips `cmd.exe` and the `vba` Node process entirely, going straight from Jest (in-process `lib/vbapm.js`) → `powershell.exe`.

## The COM bridge

`vba run` → run.ts → spawns `powershell.exe -File run-scripts/run.ps1` (or session.ps1 for the persistent mode) → run.ps1 drives Excel via COM:

1. Attach to a running visible instance (`GetActiveObject`) **or** create a new instance (hidden if `VBA_BACKGROUND_BUILD=1`).
2. Open the target file (a `.xlam` add-in, or a workbook).
3. `Application.Run("Build.ImportGraph" | "Build.CreateDocument" | "Build.ExportTo")` to relay the build/extract instruction.
4. Close the workbook (unless `keepOpen`), quit Excel (unless it was already running).

## What we just observed and fixed

- **Before the fix**: my `VBA_PERSISTENT_SESSION=1` + `.xlam` add-in changes caused (a) 20 leaked Excel instances (one per `execute()` CLI process) and (b) `Cannot run the macro 'Build.ImportGraph'` failures.
- **After reverting** those: **26/26 tests, 18/18 snapshots pass** (~280s including the rebuild).

## Current correct behavior

- The `build:check` auto-rebuilds lib when stale.
- Each test runs in a temp dir under .tmp (cleaned up unless `KEEP_E2E_TMP=1`).
- Excel instances are created/quitted per macro run (no persistence), and the coordination registry (`%TEMP%\Excel-Instances\instances.json`) tracks instances for cross-agent awareness — with `comReachable`, addins, `windowTitle`, and `null` workbooks for non-COM-reachable ones.
