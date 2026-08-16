import { copy, ensureDirSync, readFile, remove } from "fs-extra";
import { promisify } from "util";
import { run as _run } from "vbapm";
import walkSync from "walk-sync";
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
