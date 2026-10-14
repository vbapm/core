import dedent from "@timhall/dedent";
import { spawn } from "child_process";
import open from "open";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { loadProject } from "../project";
import { getTarget } from "../targets";
import { pathExists } from "../utils/fs";
import { join } from "../utils/path";

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
