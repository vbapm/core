# Background Configuration And CLI Options

## Goal

Remove the application-level dependency on `VBA_BACKGROUND_BUILD`. Background
mode should be determined by global or project-local `vba.toml` configuration,
with a command-line override for individual commands.

## Precedence: Flag > Local > Global

Resolve background mode in this order:

1. Explicit CLI flag (`--background` or `--background=false`).
2. Project-local `vba.toml`.
3. Global `vba.toml`.
4. Foreground mode by default.

The required priority is:

```text
CLI flag > local vba.toml > global vba.toml > foreground default
```

Expected examples:

```powershell
vba build --background
vba build --background=false
vba config background true
vba config --global background false
```

The positive `--background` flag is the required interface. Supporting
`--background=false` is useful for overriding a configured background default;
either explicit CLI value must override both local and global configuration.
an optional `--foreground` alias can be added if the argument parser supports
it consistently.

## Configuration Files

Global settings live next to the installed CLI:

```text
%APPDATA%\vbapm\bin\vba.toml
```

Project settings live in `vba.toml` next to `vbaproject.toml`. Local settings
override global settings, and either setting is overridden by an explicit CLI
flag. The initial supported setting is:

```toml
background = true
```

The existing `loadToolSettings`, `saveToolSettings`,
`loadEffectiveToolSettings`, and `resolveBackgroundMode` APIs should remain the
single source of truth for resolving these values.

## Implementation

### 1. Thread background mode through the execution API

- Add `background?: boolean` to `RunOptions`.
- Keep `keepOpen` independent from `background`.
- Add `background?: boolean` to add-in operation options where Excel is invoked.
- Pass the resolved value through `buildProject`, `updateProject`, export,
	extract, and run-macro call paths into `run()`.
- Do not use `background` as an implicit alias for `keepOpen`.

### 2. Update the PowerShell bridge

- Add a `-Background` switch or explicit boolean parameter to `run.ps1`.
- Use that parameter to decide whether to attach to an existing visible Excel
	instance or create a hidden instance.
- Keep background mode explicit for every invocation rather than reading a
	process-wide environment variable.
- Add `background` to persistent-session request JSON.
- Make `session.ps1` use the request value for each request.
- If a persistent session receives a mode different from its current mode,
	reset the session or reject the request clearly.
- Preserve the existing lock, cleanup, and registry behavior.

### 3. Update CLI commands

- `vbapm build` should resolve with `resolveBackgroundMode(args.background)`.
- Ensure the resolved mode reaches the Excel bridge through `BuildOptions`.
- Add the option to other commands that directly invoke Excel where appropriate.
- Keep `--open` behavior explicit: background builds may reopen the result in a
	visible user Excel instance after the hidden build completes.
- Remove any code that maps the background setting to `keepOpen`.
- Register the `config` command and document local/global settings.

### 4. Remove environment-variable ownership

- Remove `VBA_BACKGROUND_BUILD=0/1` from package scripts once the new path works.
- Remove direct `VBA_BACKGROUND_BUILD` checks from `run.ts`, `run.ps1`, and
	`session.ps1`.
- Remove hard-coded test assignments such as the peer-host override where the
	test configuration can provide the desired mode.
- Update contributor and e2e documentation.
- If backward compatibility is required, support the environment variable only
	as a deprecated fallback and emit a warning; do not let it override an
	explicit CLI or config value.

## E2E Configuration

Use a temporary project or global `vba.toml` for e2e runs instead of setting the
application environment variable. Keep these concerns separate:

- `background`: visible versus hidden Excel.
- `E2E_IN_PROCESS`: whether test commands run inside the Jest process.
- `VBA_PERSISTENT_SESSION`: whether PowerShell and Excel are reused.
- Jest worker count: test concurrency.

The peer-reference test should use the same configuration mechanism as the
other e2e tests, unless it specifically needs a documented per-test override.

## Tests

Add focused tests for:

- Global `background = true` enabling background mode.
- Local settings overriding global settings.
- Explicit CLI flags overriding both local and global settings.
- Missing settings defaulting to foreground mode.
- `--background` overriding `background = false`.
- `--background=false` overriding `background = true`.
- Explicit background values being passed to `run()`.
- PowerShell command construction including `-Background` only when expected.
- Persistent-session requests carrying the resolved background value.
- Persistent-session reset or rejection when the requested mode changes.

Retain the existing e2e coverage for visible and background workflows. Verify
that snapshots do not change solely because the mode is now config-driven.

## Rollout Sequence

1. Complete and test the config API and precedence rules.
2. Add explicit background propagation through the TypeScript execution path.
3. Update `run.ps1` and `session.ps1` to consume explicit request parameters.
4. Migrate e2e scripts and test setup away from `VBA_BACKGROUND_BUILD`.
5. Run unit tests, typecheck, format, and visible/background e2e suites.
6. Remove the compatibility environment variable or mark it deprecated.
7. Update documentation and changelog entries as needed.

## Acceptance Criteria

- `vba.toml` global and local settings control default background mode.
- Local configuration overrides global configuration.
- An explicit CLI value overrides both configuration levels.
- No normal application execution path requires `VBA_BACKGROUND_BUILD`.
- Background and foreground modes work for build, update, export/extract, and
	macro execution.
- Persistent sessions honor the requested mode and cannot mix modes silently.
- Existing visible Excel instances are not affected by background commands.
- All relevant unit and e2e tests pass without snapshot churn.
