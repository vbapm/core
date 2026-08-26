import { Args } from "mri";
import time from "pretty-hrtime";
import { buildProject } from "../actions/build-project";
import { openTarget } from "../actions/open-target";
import { updateProject } from "../actions/update-project";

const help = [
	"Build project from manifest (after backing up any existing built targets).",
	"",
	"Usage: vbapm build [options]",
	"",
	"Options:",
	"  --target=TYPE   Build target of type TYPE",
	"  --release       Exclude dev-* items from build",
	"  --vba-only      Update VBA in an existing built target",
	"  --open          Open built target"
].join("\n");

export default async function (args: Args) {
	if (args.help) {
		console.log(help);
		return;
	}

	const start = process.hrtime();
	const target = <string | undefined>args.target;
	const addin = <string | undefined>args.addin;
	const release = !!args.release;
	const vbaOnly = !!args["vba-only"];

	if (vbaOnly) {
		await updateProject({ target, addin, release, open: !!args.open });
		console.log(`Done. ${time(process.hrtime(start))}`);
		return;
	}

	const path = await buildProject({ target, addin, release });
	console.log(`Done. ${time(process.hrtime(start))}`);

	if (args.open) {
		console.log(`Opening built target: ${path}`);
		await openTarget(path);
	}
}
