import dedent from "@timhall/dedent";
import { Args } from "mri";
import time from "pretty-hrtime";
import { syncProject } from "../actions/sync-project";

const help = dedent`
  Sync VBA source into a built target (including one currently open in Excel).

  Usage: vbapm sync [options]

  Options:
    --target=TYPE   Sync VBA to a target of type TYPE
    --release       Exclude dev-* items from sync
    --open          Leave the target open in Excel after syncing`;

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

	await syncProject({ target, addin, release, open });
	console.log(`Done. ${time(process.hrtime(start))}`);
}
