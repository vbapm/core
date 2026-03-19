/**
 * pack.js — Developer utility script for packing a vbapm/vba-blocks package into a .block archive.
 *
 * Usage:
 *   node scripts/pack <path/to/package/dir> [--force]
 *
 * Arguments:
 *   dir     Path to the package directory containing a vbaproject.toml with a [package] section.
 *
 * Options:
 *   --force   Overwrite an existing .block file if one already exists.
 *
 * Output:
 *   Creates <dir>/build/<sanitized-name>-v<version>.block
 *   e.g. node scripts/pack ./addins/vba-filesystem
 *        → addins/vba-filesystem/build/vba-filesystem-v1.0.0.block
 *
 * What gets included in the archive:
 *   - vbaproject.toml (the package manifest)
 *   - All VBA source files listed under [src] in the manifest
 *   - README / CHANGELOG / LICENSE / NOTICE files if present
 *
 * Notes:
 *   - Package names containing "/" (scoped packages) are sanitized for use as filenames:
 *     "/" is replaced with "--", other unsafe characters are replaced with "-".
 *   - This script is also called internally by publish.js when a .block does not yet exist.
 */

const { dirname, resolve, relative, join } = require("path");
const { ensureDir, pathExists, readFile, remove } = require("fs-extra");
const mri = require("mri");
const ls = require("./lib/ls");
const zip = require("./lib/zip");
const sanitizeName = require("./lib/sanitize-name");

const IS_MANIFEST = /vbaproject\.toml/;
const IS_README = /readme/i;
const IS_CHANGELOG = /changes|changelog|history/i;
const IS_LICENSE = /license|licence/i;
const IS_NOTICE = /notice/i;

main().catch(error => {
	console.error(error);
	process.exit(1);
});

// Usage: node scripts/pack ./input/dir
async function main() {
	const { parse } = await import("@decimalturn/toml-patch");
	const {
		_: [input],
		force = false
	} = mri(process.argv.slice(2));

	const dir = resolve(input);
	if (!(await pathExists(dir))) {
		throw new Error(`Input directory "${input}" not found`);
	}

	const manifest_path = join(dir, "vbaproject.toml");
	if (!(await pathExists(manifest_path))) {
		throw new Error(`vbaproject.toml not found in input directory "${input}"`);
	}

	const manifest = parse(await readFile(manifest_path, "utf8"));
	if (!manifest.package) {
		throw new Error(`pack only supports packages ([package] in vbaproject.toml)`);
	}

	const { name, version } = manifest.package;
	const block_name = `${sanitizeName(name)}-v${version}.block`;
	const block_path = join(dir, "build", block_name);
	if (await pathExists(block_path)) {
		if (!force) {
			throw new Error(`A block named "${block_name}" already exists. Use --force to overwrite it`);
		} else {
			await remove(block_path);
		}
	}

	const src_files = Object.values(manifest.src).map(src => {
		return join(dir, typeof src === "string" ? src : src.path);
	});

	const files = ls(dir)
		.filter(file => {
			return (
				IS_MANIFEST.test(file) ||
				IS_README.test(file) ||
				IS_CHANGELOG.test(file) ||
				IS_LICENSE.test(file) ||
				IS_NOTICE.test(file) ||
				src_files.includes(file)
			);
		})
		.reduce((memo, file) => {
			memo[file] = relative(dir, file);
			return memo;
		}, {});

	await ensureDir(dirname(block_path));
	await zip(files, block_path);

	console.log(`Done. Created ${block_path}`);
}
