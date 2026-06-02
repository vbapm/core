import { addins, extensionToApplication } from "../addin";
import { CliError, ErrorCode } from "../errors";
import { loadProject } from "../project";
import { getTarget } from "../targets";
import { pathExists } from "../utils/fs";
import { join, resolve } from "../utils/path";
import { run } from "../utils/run";
import { isTargetSaved } from "./check-saved";

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
	const addin = addins[application];
	const resolvedFile = resolve(builtFile);
	const save = !!options.save;
	const force = !!options.force;

	if (!(await pathExists(builtFile))) {
		throw new CliError(ErrorCode.OpenTargetNotBuilt, `Built target not found at "${builtFile}".`);
	}

	if (!(await pathExists(addin))) {
		throw new CliError(
			ErrorCode.RunScriptNotFound,
			`vbapm addin not found at "${addin}". This is a fatal error and will require vbapm to be re-installed.`
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

	await run(application, addin, "Build.CloseFile", [
		JSON.stringify({
			file: resolvedFile,
			save
		})
	]);

	return builtFile;
}
