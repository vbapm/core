/**
 * publish.js — Utility script for repository owner to publishing a VBA package to the vbapm/vba-blocks registry.
 *
 * Usage:
 *   npx scripts/publish <path/to/package/dir> [--dryrun]
 *
 * What it does:
 *   1. Reads the package manifest (vbaproject.toml) from the given directory.
 *   2. Packs the package into a .block archive (if not already built).
 *   3. Uploads the .block file to the S3 package store (packages.vba-blocks.com).
 *   4. Validates the upload by downloading and comparing checksums.
 *   5. Clones or pulls the registry git repository.
 *   6. Appends a new JSON entry (name, version, deps, checksum) to the package's registry file.
 *   7. Commits and pushes the registry change.
 *
 * Options:
 *   --dryrun   Log all steps without actually uploading, writing, or pushing anything.
 *
 * Prerequisites:
 *   - AWS credentials configured for S3 upload (via .env or environment variables).
 *   - Git credentials with push access to the registry repository.
 */

import { promisify } from "node:util";
import { resolve, join } from "node:path";
import { fileURLToPath } from "node:url";
import assert from "node:assert";
import { execFile as execFileCallback } from "node:child_process";
import fsExtra from "fs-extra";
import mri from "mri";
import tmp from "tmp";
import joinUrl from "url-join";
import { parse } from "@decimalturn/toml-patch";
import checksum from "./lib/checksum.js";
import download from "./lib/download.js";
import git from "./lib/git.js";
import s3 from "./lib/s3.js";
import sanitizeName from "./lib/sanitize-name.js";

const { readFile, ensureFile, pathExists, writeFile } = fsExtra;
const execFile = promisify(execFileCallback);
const tmpDir = promisify(tmp.dir);

const scriptDirectory = fileURLToPath(new URL(".", import.meta.url));
const registry = join(scriptDirectory, "registry");

main().catch(error => {
	console.error(error);
	process.exit(1);
});

// Usage: node scripts/publish ./input/dir
async function main() {
	let {
		_: [input],
		dryrun = false
	} = mri(process.argv.slice(2));

	const dir = resolve(input);
	assert(await pathExists(dir), `Input directory "${input}" not found`);

	const manifest_path = join(dir, "vbaproject.toml");
	assert(
		await pathExists(manifest_path),
		`vbaproject.toml not found in input directory "${input}"`
	);

	const manifest = parse(await readFile(manifest_path, "utf8"));
	assert(manifest.package, `publish only supports packages ([package] in vbaproject.toml)`);

	const cksum = await uploadAndValidate(dir, manifest, { dryrun });
	const entry = generateEntry(manifest, cksum);
	await publishToRegistry(entry, { dryrun });
}

async function uploadAndValidate(dir, manifest, { dryrun }) {
	const { name, version } = manifest.package;
	const block_name = `${sanitizeName(name)}-v${version}.block`;
	const block_path = join(dir, "build", block_name);

	if (!(await pathExists(block_path))) {
		await act(dryrun, `Packing ${block_name}...`, async () => {
			await execFile("node", [join(scriptDirectory, "pack.js"), dir]);
		});
	}

	await act(dryrun, `Uploading ${block_name}...`, () =>
		s3.upload({ filename: block_name, path: block_path })
	);

	const cksum = await checksum(block_path);
	await act(dryrun, "Validating checksum...", async () => {
		const url = joinUrl("https://packages.vba-blocks.com", block_name);
		const dest = join(await tmpDir(), block_name);

		await download(url, dest);

		const comparison = await checksum(dest);
		assert(
			comparison === cksum,
			`Uploaded checksum does not match (downloaded: ${comparison} vs expected: ${cksum})`
		);
	});

	return cksum;
}

function generateEntry(manifest, cksum) {
	const { name, version } = manifest.package;
	const deps = [];

	if (manifest.dependencies) {
		for (const [name, value] of Object.entries(manifest.dependencies)) {
			if (typeof value === "string") {
				deps.push({ name, req: value });
			} else if (value.version) {
				deps.push({ name, req: value.version });
			} else {
				throw new Error(
					`Only registry dependencies are supported when publishing, "${name}" is using an supported format`
				);
			}
		}
	}

	return { name, vers: version, deps, cksum };
}

async function publishToRegistry(entry, { dryrun }) {
	const { name, vers: version } = entry;

	console.log(`${dryrun ? "[dryrun] " : ""}Publishing\n\n${JSON.stringify(entry, null, 2)}\n`);

	// 1. git pull (or git clone if not found)
	const registry_exists = await pathExists(join(registry, ".git"));
	if (!registry_exists) {
		await act(dryrun, `git clone...`, () => git.clone({ dir: scriptDirectory }));
	} else {
		await act(dryrun, `git pull "${registry}"...`, () => git.pull({ dir: registry }));
	}

	// 2. Load existing entries
	const entries_path = join(registry, sanitizeName(name));
	await act(dryrun, () => ensureFile(entries_path));

	let raw_entries = "";
	try {
		raw_entries = await readFile(entries_path, "utf8");
	} catch {}

	const entries = raw_entries
		.split("\n")
		.filter(Boolean)
		.map(entry => JSON.parse(entry));

	// 3. Add to entries
	for (const existing of entries) {
		if (existing.vers === version) {
			throw new Error(`${name} v${version} has already been published`);
		}
	}
	const added = entries.concat(entry).map(value => JSON.stringify(value));

	// 4. Write to registry
	const raw_added = `${added.join("\n")}\n`;
	await act(dryrun, `Writing to ${entries_path}...`, () => writeFile(entries_path, raw_added));

	// 5. Commit
	const message = `Publish ${name} v${version}`;
	await act(dryrun, `git commit -am "${message}"...`, () => git.commit({ dir: registry, message }));

	// 6. Push
	await act(dryrun, `git push...`, async () => {
		await git.push({ dir: registry });
	});

	console.log(`Published ${name} v${version}!`);
}

async function act(dryrun, message, action) {
	if (typeof message === "function") {
		action = message;
		message = null;
	}

	if (message) console.log(`${dryrun ? "[dryrun] " : ""}${message}`);
	if (!dryrun) await action();
}
