import { exec as _exec } from "child_process";
import { promisify } from "util";
import { extensionToApplication } from "../addin";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { loadProject } from "../project";
import { getTarget } from "../targets";
import { pathExists } from "../utils/fs";
import { join, resolve } from "../utils/path";
import { isTargetSaved } from "./check-saved";

const exec = promisify(_exec);

export interface CloseOptions {
	target?: string;
	save?: boolean;
	force?: boolean;
}

export async function closeTarget(options: CloseOptions = {}): Promise<string> {
	const project = await loadProject();
	const { target } = getTarget(project, options.target);

	const builtFile = join(project.paths.build, target.filename);
	const application = extensionToApplication(target.type);
	const resolvedFile = resolve(builtFile);
	const save = !!options.save;
	const force = !!options.force;

	const script = join(env.scripts, env.isWindows ? "run.ps1" : "run.applescript");

	if (!(await pathExists(builtFile))) {
		throw new CliError(
			ErrorCode.OpenTargetNotBuilt,
			`Built target not found at "${builtFile}".`
		);
	}

	if (!(await pathExists(script))) {
		throw new CliError(
			ErrorCode.RunScriptNotFound,
			`Bridge script not found at "${script}". This is a fatal error and will require vbapm to be re-installed.`
		);
	}

	// If neither --save nor --force, check whether the workbook has unsaved changes
	if (!save && !force) {
		const saved = await isTargetSaved(target, project);
		if (saved === false) {
			throw new CliError(
				ErrorCode.CloseTargetUnsavedChanges,
				`The workbook "${target.filename}" has unsaved changes.\n\nUse --save to save before closing, or --force to discard changes.`
			);
		}
	}

	let command: string;
	if (env.isWindows) {
		const saveFlag = save ? "-Save" : "";
		command = `powershell -NoProfile -ExecutionPolicy Bypass -File "${script}" -Close ${saveFlag} "${application}" "${resolvedFile}"`;
	} else {
		const saveArg = save ? "1" : "0";
		command = `osascript '${script}' '${application}' '${resolvedFile}' 'close' '${saveArg}'`;
	}

	await exec(command, { env: process.env });

	return builtFile;
}
