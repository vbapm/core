import dedent from "@timhall/dedent";
import { Args } from "mri";
import time from "pretty-hrtime";
import { updateProject } from "../actions/update-project";
import { resolveBackgroundMode } from "../config";

const help = dedent`
  Update VBA source in a built target (including one currently open in Excel).

  Usage: vbapm update [options]

  Options:
    --target=TYPE   Update VBA in a target of type TYPE
    --release       Exclude dev-* items from update
	--open          Leave the target open in Excel after updating
	--background    Use a hidden Excel instance`;

export default async function (args: Args) {
	if (args.help) {
		console.log(help);
		return;
	}

	const start = process.hrtime();
	const target = <string | undefined>args.target;
	const addin = <string | undefined>args.addin;
	const release = !!args.release;
	const open = !!args.open;
	const background =
		typeof args.background === "boolean" ? args.background : await resolveBackgroundMode();

	await updateProject({ target, addin, release, open, background });
	console.log(`Done. ${time(process.hrtime(start))}`);
}
