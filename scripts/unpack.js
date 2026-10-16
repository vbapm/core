/**
 * unpack.js — Developer utility script for extracting a .block archive.
 *
 * Usage:
 *   node scripts/unpack <block> [dest]
 *
 * Arguments:
 *   block   Path to the .block archive to extract.
 *   dest    (Optional) Destination directory. Defaults to the archive path
 *           with its extension removed (e.g. "foo/bar.block" → "foo/bar/").
 *
 * Examples:
 *   node scripts/unpack tests/__fixtures__/.vbapm/packages/vba-blocks/json-v2.3.0.block
 *   # Extracts to: tests/__fixtures__/.vbapm/packages/vba-blocks/json-v2.3.0/
 *
 *   node scripts/unpack tests/__fixtures__/.vbapm/packages/vba-blocks/dictionary-v1.4.1.block ./output/dictionary
 *   # Extracts to: ./output/dictionary/
 */

import { resolve, dirname, basename, extname, join } from "node:path";
import mri from "mri";
import fsExtra from "fs-extra";
import decompress from "decompress";

const { ensureDir, pathExists } = fsExtra;

main().catch(error => {
	console.error(error);
	process.exit(1);
});

// Usage: node scripts/unpack block [dest]
async function main() {
	const {
		_: [input, output]
	} = mri(process.argv.slice(2));

	const block = resolve(input);
	const dest = output ? resolve(output) : removeExtension(block);

	if (!(await pathExists(block))) {
		throw new Error(`Input block "${input}" not found`);
	}

	ensureDir(dest);
	await decompress(block, dest);
}

function removeExtension(path) {
	const dir = dirname(path);
	const base = basename(path, extname(path));

	return join(dir, base);
}
