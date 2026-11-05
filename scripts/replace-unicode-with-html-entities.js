import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { TextDecoder } from "node:util";
import { createNamedHtmlEntityReplacer } from "./html-entities.js";

const entityData = JSON.parse(
	readFileSync(new URL("./data/html-entities.json", import.meta.url), "utf8")
);
export const replaceWithNamedHtmlEntities = createNamedHtmlEntityReplacer(entityData);

export function replaceFileContents(filePath) {
	const original = readFileSync(filePath);
	let text;

	try {
		text = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(original);
	} catch {
		throw new Error(`File is not valid UTF-8: ${filePath}`);
	}

	const result = replaceWithNamedHtmlEntities(text);
	if (result.text !== text) {
		writeFileSync(filePath, result.text, "utf8");
	}

	return { replacements: result.replacements, unconverted: result.unconverted };
}

function printUsage() {
	console.log("Usage: pnpm run html-entities:replace -- <utf8-file> [<utf8-file> ...]");
	console.log(
		"Replaces supported Unicode sequences with WHATWG named character references in place."
	);
	console.log("Sequences without a named reference remain unchanged and are reported.");
}

function main(paths) {
	if (paths.length === 1 && paths[0] === "--help") {
		printUsage();
		return;
	}
	if (paths.length === 0 || paths.some(path => path.startsWith("-"))) {
		printUsage();
		process.exitCode = 2;
		return;
	}

	let failed = false;
	for (const filePath of paths) {
		try {
			const { replacements, unconverted } = replaceFileContents(filePath);
			console.log(`${filePath}: replaced ${replacements} sequence(s)`);
			if (unconverted > 0) {
				console.warn(
					`${filePath}: ${unconverted} non-ASCII code point(s) have no named reference and remain unchanged`
				);
			}
		} catch (error) {
			console.error(`${filePath}: ${error.message}`);
			failed = true;
		}
	}
	if (failed) process.exitCode = 1;
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
	const args = process.argv.slice(2);
	if (args[0] === "--") args.shift();
	main(args);
}
