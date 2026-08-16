import { copy, ensureDirSync, readFile, remove } from "fs-extra";
import { promisify } from "util";
import { run as _run, closePowerShellSession } from "vbapm";
import walkSync from "walk-sync";
import mri from "mri";
import { env } from "../../src/env";
import { tmpFolder } from "../../src/utils/fs";
import { basename, extname, join, resolve } from "../../src/utils/path";
import { RunResult } from "../../src/utils/run";
import { truncate } from "../../src/utils/text";
const exec = promisify(require("child_process").exec);

export { RunResult };

/**
 * Shared test utilities for the end-to-end (`e2e`) suite.
 *
 * These helpers orchestrate the full CLI → Excel COM pipeline against a real
 * Excel instance, and provide two distinct ways to drive `vba`:
 *
 * ## Execution paths
 *
 * 1. **`execute(cwd, command)`** — runs the `vba` CLI as a **fresh child
 *    process**. On Windows this is `child_process.exec`, which routes through
 *    `cmd.exe` (resolving the extensionless `vba` path to `bin/vba.cmd` via
 *    `PATHEXT`), which in turn launches `node.exe` running `lib/vbapm.js`.
 *    Used for `build`, `export`, `extract`, `new`, `close`, and other commands.
 *
 *    ```text
 *    Jest → child_process.exec → cmd.exe → vba.cmd → node.exe (lib/vbapm.js)
 *                                    └─ (only when a macro runs) → powershell.exe → EXCEL.EXE
 *    ```
 *
 * 2. **`run(application, file, macro, args)`** — calls the `run()` function
 *    **in-process** from the `"vbapm"` library (imported from `lib/index.js`,
 *    mapped via the e2e jest config). No intermediate `vba`/`cmd.exe` process;
 *    it goes straight from the Jest process → `spawn powershell.exe` → Excel.
 *    Used to actually invoke a macro inside a built workbook.
 *
 * ## Binary resolution (`getVbaBin`)
 *
 * `execute()` resolves the `vba` binary via `VBA_BIN_DIR` when set (to exercise
 * an **installed** CLI, e.g. `%APPDATA%\vbapm\bin\vba` from `devinstall.ps1`),
 * falling back to the checkout's `bin/vba` shim otherwise.
 *
 * ## Shared helpers
 *
 * - `tmp(id, action)` — creates a temp dir under `tests/.tmp/`, runs `action`,
 *   then removes it (unless `KEEP_E2E_TMP=1`).
 * - `setup(dir, id, action)` — copies a fixture `dir` into a temp dir, then
 *   runs `action(path)`.
 * - `readdir(cwd)` — snapshots the (non-binary, VCS-ignoring) files under a
 *   project dir as a `{ path: contents }` map for Jest snapshots.
 *
 * ## Environment variables
 *
 * - `KEEP_E2E_TMP` — keep the `tests/.tmp/` folders after a run for inspection.
 * - `E2E_VERBOSE` (or `--verbose`) — echo every executed command + its output.
 * - `VBA_BIN_DIR` — resolve the CLI from this dir instead of `bin/vba`.
 */
const tmp_dir = join(__dirname, "../.tmp");
ensureDirSync(tmp_dir);
// To keep the tmp folder around for inspection, run `$env:KEEP_E2E_TMP=1` in PowerShell or `export KEEP_E2E_TMP=1` in bash before running the tests. The tmp folder is located at `tests/.tmp`.
const keepTmp = /^(1|true|yes)$/i.test(process.env.KEEP_E2E_TMP || "");
// To enable verbose logging of executed commands and their output, set the environment variable `E2E_VERBOSE` to `1`, `true`, or `yes`.
const hasVerboseArg = process.argv.some(arg => arg === "--verbose" || arg === "-v");
const isVerbose = /^(1|true|yes)$/i.test(process.env.E2E_VERBOSE || "") || hasVerboseArg;

export async function tmp(id: string, action: (cwd: string) => void) {
	const path = await tmpFolder({ dir: tmp_dir, prefix: `${id}-` });

	try {
		await action(path);
	} finally {
		if (!keepTmp) {
			await remove(path);
		}
	}
}
/**
 * Sets up a temporary directory with the contents of `dir`, then runs `action` with the temporary directory as the current working directory. The temporary directory is removed after `action` completes, unless the environment variable `KEEP_E2E_TMP` is set to `1`, `true`, or `yes`.
 * @param dir The directory to copy into the temporary directory.
 * @param id A unique identifier for the temporary directory.
 * @param action The action to run with the temporary directory as the current working directory.
 */
export async function setup(dir: string, id: string, action: (cwd: string) => void): Promise<void> {
	await tmp(id, async path => {
		await copy(dir, path);
		await action(path);
	});
}

/**
 * Resolves the `vba` binary path.
 *
 * When the `VBA_BIN_DIR` environment variable is set, the binary is resolved
 * from that directory (e.g. `%APPDATA%\vbapm\bin`). This allows e2e tests to
 * exercise the **installed** CLI rather than the checkout's dev copy.
 *
 * Falls back to `<checkout>/bin/vba` when `VBA_BIN_DIR` is not set.
 */
function getVbaBin(binDir?: string): string {
	const dir = binDir ?? process.env.VBA_BIN_DIR;
	if (dir) {
		return resolve(dir, "vba");
	}
	return resolve(__dirname, "../../bin/vba");
}

export async function execute(
	cwd: string,
	command: string,
	options?: { binDir?: string }
): Promise<{ stdout: string; stderr: string }> {
	// Opt-in: run the command in-process (skip cmd.exe → vba.cmd → node.exe),
	// importing the action functions directly. Shares the Jest process's module
	// state (and, if a persistent session is active, the same Excel instance).
	if (/^(1|true|yes)$/i.test(process.env.E2E_IN_PROCESS || "")) {
		return executeInProcess(cwd, command);
	}

	const bin = getVbaBin(options?.binDir);
	const result = await exec(`"${bin}" ${command}`, { cwd, env: process.env });

	if (isVerbose) {
		const title = `[e2e] ${command} (${cwd})`;
		process.stdout.write(`${"=".repeat(12)} ${title} ${"=".repeat(12)}\n`);
		if (result.stdout?.length) process.stdout.write(result.stdout);
		if (result.stderr?.length) process.stderr.write(result.stderr);
		process.stdout.write(`${"=".repeat(12)} end ${title} ${"=".repeat(12)}\n`);
	}

	// Give Office time to clean up
	await wait(500);

	return result;
}

/**
 * Run a CLI command string in-process by dispatching it to the corresponding
 * action function (the same functions the CLI bin commands call), capturing the
 * `console.log`/`console.error` output as `stdout`/`stderr`.
 *
 * This avoids spawning `cmd.exe`/`vba.cmd`/`node.exe` per command and lets the
 * command share the Jest process's module state (and, when the persistent
 * session is active, the same Excel instance).
 */
async function executeInProcess(
	cwd: string,
	command: string
): Promise<{ stdout: string; stderr: string }> {
	const [cmdName, ...rest] = command.trim().split(/\s+/);
	const args = mri(rest, {});

	// Capture process stdout/stderr (console.log/error) for the command's duration.
	let stdout = "";
	let stderr = "";
	const origOut = process.stdout.write;
	const origErr = process.stderr.write;
	process.stdout.write = ((chunk: any) => {
		stdout += typeof chunk === "string" ? chunk : chunk?.toString?.() ?? "";
		return true;
	}) as any;
	process.stderr.write = ((chunk: any) => {
		stderr += typeof chunk === "string" ? chunk : chunk?.toString?.() ?? "";
		return true;
	}) as any;

	const prevCwd = process.cwd();
	const prevEnvCwd = env.cwd;
	try {
		process.chdir(cwd);
		// The CLI normally spawns a fresh process whose `env.cwd` is captured at
		// module load. Running in-process means `env.cwd` (and process.cwd) must
		// be updated to the test's dir, or `loadProject()` resolves the manifest
		// from the repo root instead of the temp project dir.
		env.cwd = cwd;
		await dispatchCommand(cmdName, args);
	} catch (err: any) {
		// Mirror the CLI's error handling (vbapm.ts handleError): clean the error
		// and write "ERROR <message>" to stderr, then throw with stdout/stderr so
		// callers can read `err.stderr || err.stdout` the same way they read a
		// spawned CLI's rejection.
		const { cleanError } = await import("../../src/errors");
		const { message } = cleanError(err);
		stderr += `ERROR ${message}\n`;
		const e: any = new Error(message);
		e.stdout = stdout;
		e.stderr = stderr;
		throw e;
	} finally {
		process.stdout.write = origOut;
		process.stderr.write = origErr;
		process.chdir(prevCwd);
		env.cwd = prevEnvCwd;
	}

	if (isVerbose) {
		const title = `[e2e in-process] ${command} (${cwd})`;
		process.stdout.write(`${"=".repeat(12)} ${title} ${"=".repeat(12)}\n`);
		if (stdout.length) process.stdout.write(stdout);
		if (stderr.length) process.stderr.write(stderr);
		process.stdout.write(`${"=".repeat(12)} end ${title} ${"=".repeat(12)}\n`);
	}

	await wait(500);
	return { stdout, stderr };
}

/**
 * Map a CLI command name to its action function and run it (in-process).
 * Mirrors the bin/vbapm-*.ts command entrypoints.
 */
async function dispatchCommand(cmdName: string, args: any): Promise<void> {
	// `mri` puts positional args in `args._`; the first positional is sometimes
	// the command itself (already stripped above), so `args._` holds the rest.
	switch (cmdName) {
		case "build": {
			const { buildProject } = await import("../../src/actions/build-project");
			await buildProject({
				target: args.target,
				addin: args.addin,
				release: !!args.release
			});
			return;
		}
		case "new": {
			const { createProject } = await import("../../src/actions/create-project");
			const [name] = args._;
			await createProject({
				name,
				target: args.target,
				from: args.from,
				pkg: !!args.package,
				git: "git" in args ? !!args.git : true,
				configTemplates: "conf" in args ? !!args.conf : true
			});
			return;
		}
		case "export":
		case "extract": {
			const { exportProject } = await import("../../src/actions/export-project");
			await exportProject({
				target: args.target,
				completed: args.completed,
				addin: args.addin,
				xmlOnly: !!args["xml-only"],
				vbaOnly: !!args["vba-only"],
				skipSheetNameNormalization: !!args["skip-sheet-name-normalization"]
			});
			return;
		}
		case "update": {
			const { updateProject } = await import("../../src/actions/update-project");
			await updateProject({
				target: args.target,
				addin: args.addin,
				release: !!args.release,
				open: !!args.open
			});
			return;
		}
		case "close": {
			const { closeTarget } = await import("../../src/actions/close-target");
			await closeTarget({ target: args.target, save: !!args.save, force: !!args.force });
			return;
		}
		case "open": {
			const { openTarget, getTargetPath } = await import("../../src/actions/open-target");
			const path = await getTargetPath(args.target);
			await openTarget(path);
			return;
		}
		case "version": {
			const { incrementVersion } = await import("../../src/actions/increment-version");
			await incrementVersion(args._[0] || "patch", { preid: args.preid });
			return;
		}
		default:
			throw new Error(`executeInProcess: unsupported command "${cmdName}"`);
	}
}

const isBackup = /\.backup/;
const isGit = /\.git[/,\\]/;
const isBinary = (file: string) => [".xlsm", ".frx"].includes(extname(file));

export async function readdir(cwd: string): Promise<{ [path: string]: string }> {
	const files = walkSync(cwd, { directories: false });
	const details: { [file: string]: string } = {};
	for (const file of files) {
		if (isBackup.test(file) || isGit.test(file)) continue;

		// TEMP Need reproducible builds to compare binary results
		if (isBinary(file)) {
			details[file] = "<TODO>";
		} else {
			const data = await readFile(resolve(cwd, file), "utf8");
			details[file] = basename(file) === "vbaproject.toml" ? data : truncate(normalize(data), 200);
		}
	}

	return details;
}

export async function run(
	application: string,
	file: string,
	macro: string,
	args: string[] = []
): Promise<RunResult> {
	let result: RunResult;
	try {
		result = await _run(application, file, macro, args);

		// Give Office time to clean up
		await wait(500);
	} catch (err) {
		const cause = err as { result?: RunResult };
		if (!cause.result) throw err;
		result = cause.result;
	}

	return result;
}

async function wait(ms: number) {
	return new Promise<void>(resolve => {
		setTimeout(resolve, ms);
	});
}

function normalize(value: string): string {
	return value.replace(/\r/g, "{CR}").replace(/\n/g, "{LF}").replace(/\t/g, "{tab}");
}

/**
 * Close the process-scoped persistent PowerShell session (and the Excel it
 * owns), if one was started during the run. Call this in a `globalSetup` /
 * `afterAll` so the in-process `run()` helper doesn't leak a hidden Excel
 * instance when `VBA_PERSISTENT_SESSION=1` is enabled.
 *
 * Note: this only affects the in-process library path (the Jest process's own
 * `PowerShellSession` singleton). It does not touch any Excel instances created
 * by spawned `vba` processes via `execute()`.
 */
export async function closePersistentSession(): Promise<void> {
	await closePowerShellSession();
}
