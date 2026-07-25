import dedent from "@timhall/dedent";
import { yellowBright } from "@timhall/ansi-colors";
import { Args } from "mri";
import { loadProject } from "../project";
import { writeManifest } from "../manifest";
import { Component } from "../build/component";

const help = dedent`
  Manage the vbaproject.toml manifest.

  Usage: vbapm manifest <subcommand> [options]

  Subcommands:
    fix       Fix [src] entries whose key does not match the file's
              Attribute VB_Name.  Scans non-wildcard entries and
              renames mismatched keys to match the file.`;

export default async function (args: Args) {
	const sub = args._[0];

	if (args.help || !sub) {
		console.log(help);
		return;
	}

	switch (sub) {
		case "fix":
			await fixSrc(args);
			break;
		default:
			console.log(help);
	}
}

async function fixSrc(_args: Args) {
	const project = await loadProject();
	const src = project.manifest.src;
	let fixed = 0;

	for (let i = 0; i < src.length; i++) {
		const source = src[i];
		if (source.path.includes("*")) continue;

		const component = await Component.load(source.path, undefined as any);
		if (source.name !== component.name) {
			console.log(
				`${yellowBright("fix:")} [src] "${source.name}" → "${component.name}" (${source.path})`
			);
			src[i] = { ...source, name: component.name };
			fixed++;
		}
	}

	if (fixed > 0) {
		await writeManifest(project.manifest, project.paths.dir);
		console.log(`\nFixed ${fixed} [src] ${fixed === 1 ? "entry" : "entries"}.`);
	} else {
		console.log("All [src] entries already match their file's Attribute VB_Name.");
	}
}
