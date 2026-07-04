import dedent from "@timhall/dedent";
import { execFile as _execFile, spawn } from "child_process";
import { promisify } from "util";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { pathExists } from "./fs";
import { has } from "./has";
import { parallel } from "./parallel";
import { join } from "./path";
import { createStdoutFile } from "./stdout-file";

const execFile = promisify(_execFile);

const debug = env.debug("vbapm:run");
const SPECIAL_FILE_STDOUT = env.isWindows ? "CON" : "/dev/stdout";

export interface RunResult {
	success: boolean;
	messages: string[];
	warnings: string[];
	errors: string[];
	stdout?: string;
	stderr?: string;
}

export class RunError extends Error {
	result: RunResult;

	constructor(result: RunResult) {
		const message = result.errors.join("\n") || "An unknown error occurred.";
		super(message);

		this.result = result;
	}
}

export function isRunError(error: Error | RunError): error is RunError {
	return has(error, "result");
}

export interface RunOptions {
	keepOpen?: boolean;
}

export async function run(
	application: string,
	file: string,
	macro: string,
	args: string[],
	options: RunOptions = {}
): Promise<RunResult> {
	const script = join(env.scripts, env.isWindows ? "run.ps1" : "run.applescript");

	if (!(await pathExists(script))) {
		throw new CliError(
			ErrorCode.RunScriptNotFound,
			dedent`
        Bridge script not found at "${script}".

        This is a fatal error and will require vbapm to be re-installed.
      `
		);
	}

	const formatted_args = await parallel(args, async arg => {
		if (arg === SPECIAL_FILE_STDOUT) {
			return await createStdoutFile();
		}

		return env.isWindows ? escape(arg) : arg;
	});
	const keepOpen = !!options.keepOpen;
	// Windows uses a named switch; macOS receives keepOpen as a positional arg (position 4)
	const parts = env.isWindows
		? [application, file, macro, ...formatted_args]
		: [application, file, macro, keepOpen ? "1" : "0", ...formatted_args];
	const command = env.isWindows ? "powershell" : "osascript";
	const commandArgs = env.isWindows
		? [
				"-NoProfile",
				"-ExecutionPolicy",
				"Bypass",
				"-File",
				script,
				...(keepOpen ? ["-KeepOpen"] : []),
				...parts
			]
		: [script, ...parts];

	debug("params:", { application, file, macro, args });
	debug("command:", command, commandArgs);

	let result;
	try {
		// Use execPowershell on Windows (spawn-based) to work around Node.js libuv assertion bug
		// and execFile on macOS to avoid shell injection (no shell on either platform).
		const { stdout, stderr } = env.isWindows
			? await execPowershell(script, keepOpen, parts, { env: process.env })
			: await execFile(command, commandArgs, { env: process.env });
		result = toResult(stdout, stderr);
	} catch (err: any) {
		result = toResult(err?.stdout, err?.stderr, err);
	}

	if (!result.success) {
		throw new RunError(result);
	}

	debug("result:", result);
	return result;
}

/**
 * Workaround for Node.js v24 libuv assertion crash on Windows.
 * Uses spawn instead of exec to avoid UV_HANDLE_CLOSING race condition
 * in child_process pipe management.
 *
 * Unlike the original exec() which goes through cmd.exe, this spawns
 * powershell.exe directly with an args array to avoid shell quote issues.
 *
 * TODO: Remove this workaround once the upstream fix lands.
 *       https://github.com/nodejs/node/issues/56645
 *       Possibly a Fix PR: https://github.com/nodejs/node/pull/61999
 */
function execPowershell(
	script: string,
	keepOpen: boolean,
	parts: string[],
	options: { env: typeof process.env }
): Promise<{ stdout: string; stderr: string }> {
	return new Promise((resolve, reject) => {
		const args = [
			"-NoProfile",
			"-ExecutionPolicy",
			"Bypass",
			"-File",
			script,
			...(keepOpen ? ["-KeepOpen"] : []),
			...parts
		];

		const child = spawn("powershell.exe", args, {
			...options,
			windowsHide: true,
			stdio: ["ignore", "pipe", "pipe"]
		});

		let stdout = "";
		let stderr = "";

		child.stdout?.on("data", (data: Buffer) => {
			stdout += data.toString();
		});
		child.stderr?.on("data", (data: Buffer) => {
			stderr += data.toString();
		});

		child.on("error", err => reject(err));
		child.on("close", code => {
			if (code === 0) {
				resolve({ stdout, stderr });
			} else {
				const error = new Error(stderr || `Command failed with exit code ${code}`);
				(error as any).stdout = stdout;
				(error as any).stderr = stderr;
				reject(error);
			}
		});
	});
}

export function escape(value: string): string {
	// Replace quotes with ^q placeholder to avoid issues with shell argument passing.
	// The PowerShell/AppleScript bridge script unescapes these back to quotes.
	return value.replace(/\"/g, "^q");
}

export function unescape(value: string): string {
	return value.replace(/\^q/g, '"');
}

export function toResult(stdout: string, stderr: string, err?: Error): RunResult {
	let success = false;
	let messages: string[] = [];
	let warnings: string[] = [];
	let errors: string[] = [];

	if (stdout) {
		try {
			// For vbapm run, check for standard JSON result
			const parsed = JSON.parse(stdout);

			if ("success" in parsed && ("messages" in parsed || "errors" in parsed)) {
				({ success, messages = [], warnings = [], errors = [] } = parsed);
			} else {
				throw new Error("(ok, non-standard response)");
			}
		} catch (err) {
			success = true;
			messages = [stdout];
		}
	}

	if (err) {
		success = false;
		errors.push(unescape(err.message));
	}
	if (stderr) {
		success = false;
		errors.push(stderr);
	}

	return { success, messages, warnings, errors, stdout, stderr };
}
