import dedent from "@timhall/dedent";
import { Args } from "mri";
import time from "pretty-hrtime";
import { buildProject } from "../actions/build-project";
import { openTarget } from "../actions/open-target";
import { resolveBackgroundMode } from "../config";

const help = dedent`
  Build project from manifest (after backing up any existing built targets).

  Usage: vbapm build [options]

  Options:
    --target=TYPE   Build target of type TYPE
    --release       Exclude dev-* items from build
    --background    Run the build in background mode (hidden Excel instance)
    --open          Keep the built target open after build (opens in visible Excel instance afterward for background builds)`;

export default async function (args: Args) {
	if (args.help) {
		console.log(help);
		return;
	}

	const start = process.hrtime();
	const target = <string | undefined>args.target;
	const addin = <string | undefined>args.addin;
	const release = !!args.release;

	const background =
		typeof args.background === "boolean" ? !!args.background : await resolveBackgroundMode();

	// In a visible (foreground) build, `open` keeps the freshly imported
	// workbook open in the user's existing Excel session. In a background build
	// the hidden instance is torn down, so instead pass no `open` and reopen the
	// built file in a visible Excel afterward.
	const open = !!args.open && !background;

	const path = await buildProject({
		target,
		addin,
		release,
		open,
		background
	});
	console.log(`Done. ${time(process.hrtime(start))}`);

	if (args.open) {
		console.log(`Opening built target: ${path}`);
		if (background) {
			await openTarget(path);
		}
	}
}
