import dedent from "@timhall/dedent";
import { Args } from "mri";
import { getTargetPath, openTarget } from "../actions/open-target";

const help = dedent`
  Open the current built target file.

  Usage: vbapm open [options]

  Options:
    --target=TYPE   Open built target of type TYPE`;

export default async function (args: Args) {
	if (args.help) {
		console.log(help);
		return;
	}

	const target = <string | undefined>args.target;

	const path = await getTargetPath(target);
	console.log(`Opening built target: ${path}`);
	await openTarget(path);
}
