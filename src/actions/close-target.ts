import { exec as _exec } from "child_process";
import { promisify } from "util";
import { extensionToApplication } from "../addin";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { loadProject } from "../project";
import { getTarget } from "../targets";
import { pathExists } from "../utils/fs";
import { join, resolve } from "../utils/path";

const exec = promisify(_exec);

export interface CloseOptions {
	target?: string;
	save?: boolean;
}

export async function closeTarget(options: CloseOptions = {}): Promise<string> {
	const project = await loadProject();
	const { target } = getTarget(project, options.target);

	const builtFile = join(project.paths.build, target.filename);
	const application = extensionToApplication(target.type);
	const resolvedFile = resolve(builtFile);
	const save = !!options.save;

	if (env.isWindows) {
		const script = join(env.scripts, "close.ps1");

		if (!(await pathExists(script))) {
			throw new CliError(
				ErrorCode.CloseScriptNotFound,
				`Close script not found at "${script}". This is a fatal error and will require vbapm to be re-installed.`
			);
		}

		const saveFlag = save ? "-Save" : "";
		const command = `powershell -NoProfile -ExecutionPolicy Bypass -File "${script}" "${application}" "${resolvedFile}" ${saveFlag}`;

		try {
			await exec(command, { env: process.env });
		} catch {
			// Non-fatal: workbook may not be open or Excel may not be running
		}
	} else {
		const script = join(env.scripts, "close.applescript");

		if (!(await pathExists(script))) {
			throw new CliError(
				ErrorCode.CloseScriptNotFound,
				`Close script not found at "${script}". This is a fatal error and will require vbapm to be re-installed.`
			);
		}

		const saveArg = save ? "1" : "0";
		const command = `osascript '${script}' '${application}' '${resolvedFile}' '${saveArg}'`;

		try {
			await exec(command, { env: process.env });
		} catch {
			// Non-fatal: workbook may not be open or Excel may not be running
		}
	}

	return builtFile;
}
