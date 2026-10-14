import dedent from "@timhall/dedent";
import open from "open";
import { CliError, ErrorCode } from "../errors";
import { loadProject } from "../project";
import { getTarget } from "../targets";
import { pathExists } from "../utils/fs";
import { join } from "../utils/path";

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
		const exitCode = await new Promise<number | null>((resolve, reject) => {
			childProcess.once("error", reject);
			childProcess.once("close", resolve);
		});
		if (exitCode !== 0) {
			throw new Error(`Launcher exited with code ${exitCode}`);
		}
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
