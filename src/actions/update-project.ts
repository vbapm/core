import dedent from "@timhall/dedent";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { Message } from "../messages";
import { fetchDependencies, loadProject } from "../project";
import { importTarget } from "../targets/build-target";
import { getTarget } from "../targets/get-target";
import { pathExists } from "../utils/fs";
import { join } from "../utils/path";

export interface UpdateOptions {
	target?: string;
	addin?: string;
	release?: boolean;
	open?: boolean;
}

export async function updateProject(options: UpdateOptions = {}): Promise<string> {
	env.reporter.log(Message.UpdateProjectLoading, `[1/2] Loading project...`);

	const project = await loadProject();
	const { target } = getTarget(project, options.target);
	const dependencies = await fetchDependencies(project);

	// Guard: a built file must already exist — update does not create one
	const builtFile = join(project.paths.build, target.filename);
	if (!(await pathExists(builtFile))) {
		throw new CliError(
			ErrorCode.UpdateTargetNotBuilt,
			dedent`
        No built target found for "${target.name}" at "${builtFile}".

        Run "vbapm build" first to create the built target.
      `
		);
	}

	env.reporter.log(Message.UpdateTargetUpdating, `\n[2/2] Updating VBA in "${target.filename}"...`);

	await importTarget(target, { project, dependencies }, builtFile, {
		addin: options.addin,
		release: options.release,
		open: options.open
	});

	return builtFile;
}
