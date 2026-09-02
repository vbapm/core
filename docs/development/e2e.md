Here's a summary of the e2e testing process in this repo.

## What e2e CLI tests exercise

The end-to-end tests (excel.e2e.ts + helper in execute.ts) run the full pipeline — CLI binary → Excel COM automation via the `vbapm.xlam` add-in — against real Excel on Windows. They cover build, `export`, `extract`, `update`, `new`, `close`, `run`, and worksheet/drawing edge cases, comparing output against Jest snapshots.

## How the tests are invoked

From package.json:

| Script | What it does |
|---|---|
| `pnpm run test:e2e` | explicit foreground override, **spawns the real CLI** per command |
| `pnpm run test:e2e:in-process` | visible Excel + `E2E_IN_PROCESS=1`, **dispatches CLI commands in-process** (see below) |
| `pnpm run test:e2e:background` | explicit background override (hidden Excel) |
| `pnpm run test:e2e:session` | background + `E2E_IN_PROCESS=1` + `VBA_PERSISTENT_SESSION=1`; reuses a PowerShell/Excel session per worker |
| `pnpm run test:e2e:updateSnapshots` | same as background + `--updateSnapshot` |
| `pnpm run test:e2e:multilang` | separate multilang config |

Each script is prefixed with `pnpm run build:check` → `node scripts/ensure-fresh-build.ts`, which rebuilds lib if sources are newer than the build output (that's the `[ensure-fresh-build] lib/ is stale … rebuilding` line you saw).

For Excel ownership and teardown diagnostics, set `VBA_DEBUG_INSTANCES=1`.
Lifecycle records are written to `%TEMP%\Excel-Instances\instances.log` (or
the path in `VBA_INSTANCE_LOG`) without being mixed into test stdout/stderr.
For per-test and per-command timing in the export suites, set `E2E_TIMING=1`.

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

- Default pool size is **12**; set `VBA_EXCEL_POOL_SIZE` to a different value to
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

```text
Spawned CLI path (default)
execute() in tests/__helpers__/execute.ts
└─ child_process.exec
   └─ cmd.exe → vba.cmd → node.exe → lib/vbapm.js → powershell.exe → EXCEL.EXE

In-process path (branch feature)
execute() in tests/__helpers__/execute.ts
└─ executeInProcess()
   └─ dispatchCommand()
      └─ src/bin/vbapm-*.ts
         └─ src/utils/run.ts
            └─ powershell.exe → EXCEL.EXE
```

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

1. Attach to a running visible instance (`GetActiveObject`) **or** create a new hidden instance when background mode is selected.
2. Open the target file (a `.xlam` add-in, or a workbook).
3. `Application.Run("Build.ImportGraph" | "Build.CreateDocument" | "Build.ExportTo")` to relay the build/extract instruction.
4. Close the workbook (unless `keepOpen`), quit Excel (unless it was already running).

## What we just observed and fixed

- **Before the fix**: my `VBA_PERSISTENT_SESSION=1` + `.xlam` add-in changes caused (a) 20 leaked Excel instances (one per `execute()` CLI process) and (b) `Cannot run the macro 'Build.ImportGraph'` failures.
- **After reverting** those: **26/26 tests, 18/18 snapshots pass** (~280s including the rebuild).

## Current correct behavior

- The `build:check` auto-rebuilds lib when stale.
- Each test runs in a temp dir under .tmp (cleaned up unless `KEEP_E2E_TMP=1`).
- Default e2e mode creates and quits Excel per macro run. The opt-in session mode reuses PowerShell and Excel within a test worker, then closes the session at each temp-project boundary and worker teardown.
- The coordination registry tracks instances across agents for cross-agent awareness.

## Future work: cross-process Excel bridge daemon

The current persistent session is process-local. It reuses PowerShell and Excel
only while one Node/Jest worker remains alive. A spawned `vba` command still
creates a new Node process, so it cannot discover or reuse that session. The
longer-term solution is a Windows-only broker process that owns the Excel pool
and accepts requests from short-lived CLI processes.

### Responsibilities

The daemon should:

- Own one or more hidden `Excel.Application` instances.
- Keep the PowerShell COM bridge alive for the lifetime of each owned instance.
- Serialize requests per Excel instance while allowing independent instances to
   run concurrently.
- Load `vbapm.xlam` once per instance and clean normal workbooks after every
   request.
- Track owner PID, daemon PID, instance PID, active request, open workbooks,
   loaded add-ins, and last activity in the existing Excel instance registry.
- Quit every owned Excel instance when the daemon shuts down or loses ownership.

### IPC and lifecycle

Use a Windows named pipe rather than a TCP port. The pipe name should include
the current user and repository or environment identity, for example:
`\\.\pipe\vbapm-excel-{user}-{scope}`.

The CLI flow would be:

1. Locate the daemon endpoint from a small file in `%TEMP%\vbapm`.
2. Validate the endpoint's daemon PID and protocol version.
3. Start the daemon if the endpoint is absent or stale.
4. Send one framed request and wait for one framed response.
5. Keep the daemon alive for subsequent CLI invocations.
6. Send an explicit shutdown request when an owning test run ends.

The existing base64 JSON request shape is a suitable starting point. The daemon
protocol should add `protocolVersion`, `requestId`, `sessionKey`, `keepOpen`,
and cleanup policy fields. Responses should include `success`, `stdout`,
`stderr`, `errors`, and the selected instance identifier. Add `ping` and
owner-restricted `shutdown` requests.

Frames must be length-prefixed or use the existing record separators. JSON
should remain inside the frame so paths and macro arguments never pass through a
shell parser.

### Ownership and recovery

The daemon must have an explicit owner and lease:

- Write the daemon PID, pipe name, protocol version, and start time to the
   endpoint file using an atomic rename.
- Refresh a heartbeat or lease timestamp while serving requests.
- Treat an endpoint as stale when the PID is dead, the pipe cannot be opened, or
   the lease exceeds its timeout.
- Never let a stale client kill a daemon it does not own.
- After a daemon crash, start a replacement and let registry cleanup mark dead
   Excel processes inactive.
- On PowerShell or Excel failure, discard that instance, release COM references,
   and start a replacement before retrying once.
- Retry only requests known to be safe. Build/export requests need an
   idempotency key or a target-state check before retrying.

### Instance pool

The daemon should maintain a bounded pool rather than one global Excel object:

```text
CLI process
   | named pipe
   v
Excel bridge daemon
   +-- request queue
   +-- instance 1: PowerShell session + Excel.Application
   +-- instance 2: PowerShell session + Excel.Application
   +-- instance N: PowerShell session + Excel.Application
```

Route requests touching the same workbook or project to the same instance when
possible. Otherwise assign the least-busy healthy instance. Enforce the same
global limit as `VBA_MAX_EXCEL_INSTANCES` and expose queue depth and instance
health in debug logs.

Each request should decode and validate arguments, ensure the add-in, open or
reuse the workbook, run the macro, close request-owned workbooks, release COM
references, and return the result without quitting a healthy instance. If
cleanup cannot prove that an instance is clean, retire it instead of reusing it.

### Security and operability

- Restrict the named pipe ACL to the current user.
- Accept only the structured run request, never arbitrary PowerShell or commands.
- Limit request and argument sizes.
- Log daemon PID, client PID, request ID, instance ID, and macro name, but never
   workbook contents or secrets.
- Add `vbapm daemon status`, `vbapm daemon stop`, and a diagnostic command for
   queue and instance state.
- Keep the existing registry as an observability and cleanup side channel, not
   as the transport protocol.

### Implementation sequence

1. Extract the current PowerShell session request handler behind an interface.
2. Add a daemon executable and named-pipe server with `ping`, `run`, and
    `shutdown` messages.
3. Move instance creation, pooling, and cleanup into the daemon.
4. Change `src/utils/run.ts` to prefer the daemon when enabled and fall back to
    the one-shot bridge when it is unavailable.
5. Add crash, stale-endpoint, concurrent-client, and cleanup tests.
6. Run spawned-CLI e2e tests through the daemon and compare runtime, Excel
    launch count, and lingering-instance count against `--runInBand`.

### Acceptance criteria

Path B is ready only when:

- Spawned CLI invocations reuse daemon-owned Excel instances across commands.
- A daemon crash or stale endpoint self-heals without killing visible user Excel.
- Parallel clients stay within the configured instance limit.
- A failed request cannot poison the next request assigned to that instance.
- All daemon-owned Excel and PowerShell processes are gone after shutdown.
- Full spawned and in-process e2e suites pass without snapshot changes.
- Runtime and Excel launch counts are measured before and after the change.
