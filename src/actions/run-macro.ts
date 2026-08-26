import { extensionToApplication } from "../addin";
import { CliError, ErrorCode } from "../errors";
import { loadProject } from "../project";
import { getTarget } from "../targets";
import { resolveBackgroundMode } from "../config";
import { extname, join, resolve } from "../utils/path";
import { run, RunResult } from "../utils/run";

export interface RunOptions {
	target?: string;
	file?: string;
	macro: string;
	args: string[];
	keepOpen?: boolean;
}

export async function runMacro(options: RunOptions): Promise<RunResult> {
	let { target: targetType, file, macro, args = [""], keepOpen } = options;
	keepOpen ??= await resolveBackgroundMode();

	if (!file) {
		const project = await loadProject();
		const { target } = getTarget(project, targetType);

		file = join(project.paths.build, target.filename);
	}

	if (!file) {
		throw new CliError(
			ErrorCode.RunMissingFile,
			`file is required for vbapm run (e.g. vbapm run --file FILE <macro> <arg>...).`
		);
	}
	if (!macro) {
		throw new CliError(
			ErrorCode.RunMissingMacro,
			`macro is required for vbapm run (e.g. vbapm run --file FILE <macro> <arg>...).`
		);
	}

	const application = extensionToApplication(extname(file));
	const result = await run(application, resolve(file), macro, args, { keepOpen });
	const { stdout } = result;

	if (stdout && stdout.trim().length) console.log(stdout);

	return result;
}
