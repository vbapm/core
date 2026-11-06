import dedent from "@timhall/dedent";
import { spawn } from "child_process";
import open from "open";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { loadProject } from "../project";
import { getTarget } from "../targets";
import { pathExists } from "../utils/fs";
import { join } from "../utils/path";
import { convert as convertToToml } from "../utils/toml";

const EXCEL_OPEN_TIMEOUT_SECONDS = 30;
const EXCEL_OPEN_POLL_MILLISECONDS = 250;

/**
 * Return path to built target file.
 */
export async function getTargetPath(target?: string): Promise<string> {
	const project = await loadProject();
	const { target: resolvedTarget } = getTarget(project, target);

	const builtFile = join(project.paths.build, resolvedTarget.filename);
	if (!(await pathExists(builtFile))) {
		throw new CliError(
			ErrorCode.OpenTargetNotBuilt,
			dedent`
        No built target found for "${resolvedTarget.name}" at "${builtFile}".

        Run "vbapm build" first to create the built target.
      `
		);
	}

	return builtFile;
}

export async function openTarget(target: string): Promise<boolean> {
	try {
		const childProcess = await open(target, { wait: false });
		childProcess.ref();
		const exitCode = await waitForProcess(childProcess);
		if (exitCode !== 0) {
			throw new Error(`Launcher exited with code ${exitCode}`);
		}
		await waitForExcelWorkbook(target);
		return true;
	} catch (error) {
		const reason = error instanceof Error && error.message ? `\n\nReason: ${error.message}` : "";

		throw new CliError(
			ErrorCode.OpenTargetFailed,
			dedent`
        Failed to open target "${target}".

        Verify that the file exists and that your system has an application associated with this file type.${reason}
      `
		);
	}
}

export async function checkTargetOpen(target: string, filenameOnly = false): Promise<string> {
	const scriptName = env.isWindows
		? "check-open.ps1"
		: process.platform === "darwin"
			? "check-open.applescript"
			: undefined;
	if (!scriptName) return formatOpenCheckReport(target, [], [], filenameOnly);

	const script = join(env.scripts, scriptName);
	if (!(await pathExists(script))) {
		throw new Error(`Excel open-check script not found at "${script}".`);
	}

	const command = env.isWindows ? "powershell.exe" : "osascript";
	const args = env.isWindows
		? [
				"-NoProfile",
				"-ExecutionPolicy",
				"Bypass",
				"-File",
				script,
				"-Path",
				target,
				...(filenameOnly ? ["-NameOnly"] : [])
			]
		: [script, target, ...(filenameOnly ? ["filename-only"] : [])];

	return new Promise((resolve, reject) => {
		const child = spawn(command, args, {
			windowsHide: true,
			stdio: ["ignore", "pipe", "pipe"]
		});

		let stdout = "";
		let stderr = "";
		child.stdout?.on("data", data => {
			stdout += data.toString();
		});
		child.stderr?.on("data", data => {
			stderr += data.toString();
		});
		child.on("error", reject);
		child.on("close", code => {
			if (code !== 0) {
				reject(new Error(stderr.trim() || `Excel open check exited with code ${code}.`));
				return;
			}

			try {
				const result = JSON.parse(stdout.trim()) as {
					instances?: OpenCheckInstance[];
					uninspected_excel_process_ids?: number[];
				};
				if (!Array.isArray(result.instances)) {
					throw new Error("Open-check result is missing its instances array.");
				}
				void formatOpenCheckReport(
					target,
					result.instances,
					result.uninspected_excel_process_ids ?? [],
					filenameOnly
				).then(resolve, reject);
			} catch (error) {
				reject(
					new Error(
						`Unexpected Excel open-check result: ${error instanceof Error ? error.message : String(error)}`
					)
				);
			}
		});
	});
}

interface OpenCheckInstance {
	path: string;
	same_path: boolean | null;
	source: string;
	visible: boolean;
	process_id?: number | null;
	window_handle?: number | null;
	access?: "read_only" | "read_write" | "unknown";
}

async function formatOpenCheckReport(
	target: string,
	instances: OpenCheckInstance[],
	uninspectedExcelProcessIds: number[],
	filenameOnly: boolean
): Promise<string> {
	const sameName = filenameOnly ? target : (target.split(/[\\/]/).pop() ?? target);
	const reportInstances = instances.map(instance => ({
		file_path: instance.path,
		match: filenameOnly
			? "filename_only"
			: instance.same_path
				? "exact_path"
				: "same_name_different_location",
		source: instance.source,
		visible: instance.visible,
		...(instance.process_id != null ? { process_id: instance.process_id } : {}),
		...(instance.window_handle != null ? { window_handle: instance.window_handle } : {}),
		access: instance.access ?? "unknown"
	}));

	return convertToToml({
		target: {
			...(filenameOnly ? {} : { path: target }),
			name: sameName,
			open: filenameOnly ? instances.length > 0 : instances.some(instance => instance.same_path),
			...(filenameOnly
				? {}
				: { same_name_elsewhere: instances.some(instance => !instance.same_path) }),
			match_count: instances.length,
			...(uninspectedExcelProcessIds.length > 0
				? { uninspected_excel_process_ids: uninspectedExcelProcessIds }
				: {})
		},
		instances: reportInstances
	});
}

function waitForProcess(childProcess: ReturnType<typeof spawn>): Promise<number | null> {
	return new Promise((resolve, reject) => {
		childProcess.once("error", reject);
		childProcess.once("close", resolve);
	});
}

async function waitForExcelWorkbook(target: string): Promise<void> {
	if (!env.isWindows) return Promise.resolve();

	const script = join(env.scripts, "wait-open.ps1");
	if (!(await pathExists(script))) {
		throw new Error(`Readiness script not found at "${script}".`);
	}

	return new Promise((resolve, reject) => {
		const child = spawn(
			"powershell.exe",
			[
				"-NoProfile",
				"-ExecutionPolicy",
				"Bypass",
				"-File",
				script,
				"-Path",
				target,
				"-TimeoutSeconds",
				String(EXCEL_OPEN_TIMEOUT_SECONDS),
				"-PollMilliseconds",
				String(EXCEL_OPEN_POLL_MILLISECONDS)
			],
			{ windowsHide: true, stdio: ["ignore", "pipe", "pipe"] }
		);

		let stderr = "";
		child.stderr?.on("data", data => {
			stderr += data.toString();
		});
		child.on("error", reject);
		child.on("close", code => {
			if (code === 0) {
				resolve();
				return;
			}

			reject(new Error(stderr.trim() || `Excel readiness check exited with code ${code}.`));
		});
	});
}
