import dedent from "@timhall/dedent";
import { Args } from "mri";
import { addSource } from "../actions/add-source";

const help = dedent`
  Create and register a new source file in vbaproject.toml

  Usage: vbapm add <name> [options]

  Options:
    <name>          Source name (optionally with .bas or .cls extension)
    --type=TYPE     Source type (module or class)
    --dev           Add to [dev-src] instead of [src]

  Examples:
  vbapm add Module1
  vbapm add JsonParser --type class
  vbapm add TestHelpers --dev
  `;

export default async function (args: Args) {
	if (args.help) {
		console.log(help);
		return;
	}

	const [name] = args._;
	const type = <string | undefined>args.type;
	const dev = !!args.dev;

	const result = await addSource({ name, type, dev });
	if (result.isNew) {
		console.log(`Created ${result.path}`);
	} else {
		console.log(`Registered existing file ${result.path}`);
	}
}
