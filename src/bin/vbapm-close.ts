import dedent from "@timhall/dedent";
import { Args } from "mri";
import { closeTarget } from "../actions/close-target";

const help = dedent`
  Close the current built target file.

  Usage: vbapm close [options]

  Options:
    --target=TYPE   Close built target of type TYPE
    --save          Save changes before closing (default: discard)`;

export default async function (args: Args) {
	if (args.help) {
		console.log(help);
		return;
	}

	const target = <string | undefined>args.target;
	const save = !!args.save;

	const path = await closeTarget({ target, save });
	console.log(`Closed built target: ${path}`);
}
