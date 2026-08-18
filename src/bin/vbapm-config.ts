import dedent from "@timhall/dedent";
import { Args } from "mri";
import { loadToolSettings, saveToolSettings } from "../config";

const help = dedent`
  Get and set vbapm configuration values.

  Usage: vbapm config [--global] [key] [value]

  Commands:
    vbapm config                  Show the local configuration
    vbapm config --global        Show the global configuration
    vbapm config background true
    vbapm config --global background false

  Options:
    --global     Edit the global config file next to the executable
    --list       Show all config keys
    --unset      Remove a key from the active config`;

export default async function (args: Args) {
	if (args.help) {
		console.log(help);
		return;
	}

	const useGlobal = !!args.global;
	const key = Array.isArray(args._) ? args._[0] : undefined;
	const value = Array.isArray(args._) ? args._[1] : undefined;

	if (!key && !value && !args.list && !args.unset) {
		const settings = await loadToolSettings({ global: useGlobal });
		if (Object.keys(settings).length === 0) {
			console.log("(empty)");
			return;
		}
		for (const [name, item] of Object.entries(settings)) {
			console.log(`${name}=${String(item)}`);
		}
		return;
	}

	if (args.unset) {
		if (!key) {
			throw new Error("A key name is required when using --unset.");
		}
		const settings = await loadToolSettings({ global: useGlobal });
		delete settings[key];
		await saveToolSettings(settings, { global: useGlobal });
		return;
	}

	if (!key) {
		throw new Error("A key name is required.");
	}

	if (typeof value === "undefined") {
		const settings = await loadToolSettings({ global: useGlobal });
		console.log(String(settings[key] ?? ""));
		return;
	}

	const nextValue = value === "true" ? true : value === "false" ? false : value;
	await saveToolSettings({ [key]: nextValue } as any, { global: useGlobal });
	console.log(`${key}=${String(nextValue)}`);
}
