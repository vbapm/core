import { Args } from "mri";
import { checkTargetOpen, getTargetPath, openTarget } from "../actions/open-target";
import { basename, isAbsolute, resolve } from "../utils/path";

const help = [
	"Open an Excel file or the current built target, or report matching files in Excel.",
	"",
	"Usage: vbapm open [options]",
	"",
	"Options:",
	"  --target=TYPE   Use built target of type TYPE",
	"  --file=PATH     Open this file; with --check, non-absolute paths match by filename",
	"  --check         Print a TOML report of open Excel instances"
].join("\n");

export default async function (args: Args) {
	if (args.help) {
		console.log(help);
		return;
	}

	const target = <string | undefined>args.target;
	const file = args.file;
	if (file !== undefined && (typeof file !== "string" || file.length === 0)) {
		throw new Error("--file requires a path.");
	}

	const filenameOnly = args.check === true && file !== undefined && !isAbsolute(file);
	const path =
		file === undefined
			? await getTargetPath(target)
			: filenameOnly
				? basename(file)
				: resolve(file);
	if (args.check) {
		console.log(await checkTargetOpen(path, filenameOnly));
		return;
	}

	console.log(file === undefined ? `Opening built target: ${path}` : `Opening file: ${path}`);
	await openTarget(path);
}
