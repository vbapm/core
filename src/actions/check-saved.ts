import { exec as _exec } from "child_process";
import { promisify } from "util";
import { extensionToApplication } from "../addin";
import { env } from "../env";
import { Target } from "../manifest/target";
import { Project } from "../project";
import { pathExists } from "../utils/fs";
import { join, resolve } from "../utils/path";

const exec = promisify(_exec);

/**
 * Checks whether the built target workbook has been saved in the running Excel instance.
 *
 * Returns:
 *   true  — workbook is saved (or Excel is not running / workbook is not open)
 *   false — workbook has unsaved changes
 *   null  — could not determine (script missing, unexpected error)
 */
export async function isTargetSaved(target: Target, project: Project): Promise<boolean | null> {
	const application = extensionToApplication(target.type);
	const resolvedFile = resolve(join(project.paths.build, target.filename));
	const script = join(env.scripts, env.isWindows ? "run.ps1" : "run.applescript");

	if (!(await pathExists(script))) return null;

	let command: string;
	if (env.isWindows) {
		command = `powershell -NoProfile -ExecutionPolicy Bypass -File "${script}" -CheckSaved "${application}" "${resolvedFile}"`;
	} else {
		command = `osascript '${script}' '${application}' '${resolvedFile}' 'check-saved'`;
	}

	try {
		const { stdout } = await exec(command, { env: process.env });
		const result = JSON.parse(stdout.trim()) as { success: boolean; saved?: boolean };
		if (result.success && typeof result.saved === "boolean") {
			return result.saved;
		}
		return null;
	} catch {
		return null;
	}
}
