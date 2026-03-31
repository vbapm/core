import { readFile as _readFile, symlink as _symlink, writeFile as _writeFile } from "fs";
import { isAbsolute } from "path";
import { copy, emptyDir, ensureDir, ensureDirSync, move, pathExists, readJson, remove } from "fs-extra";
import { promisify } from "util";
import { hash } from "./hash";

const readFile = promisify(_readFile);
const symlink = promisify(_symlink);
const writeFile = promisify(_writeFile);

async function checksum(file: string, algorithm = "sha256"): Promise<string> {
	const data = await readFile(file);
	return hash(data, { algorithm });
}

export interface TmpOptions {
	dir?: string;
	prefix?: string;
	template?: string;
}

async function tmpFile(options: TmpOptions = {}): Promise<string> {
	const { env } = await import("../env");
	const { file: tmpFile } = await import("tmp");
	const { dir = env.temp, prefix, template } = options;
	const tmpOptions: {
		dir?: string;
		tmpdir?: string;
		prefix?: string;
		template?: string;
	} = { prefix, template };

	if (isAbsolute(dir)) {
		tmpOptions.tmpdir = dir;
	} else {
		tmpOptions.dir = dir;
	}

	await ensureDir(dir);
	return new Promise<string>((resolve, reject) => {
		// Defer requiring tmp as it adds process listeners that can cause warnings
		tmpFile(tmpOptions, (err: any, path: string) => {
			if (err) return reject(err);
			resolve(path);
		});
	});
}

async function tmpFolder(options: TmpOptions = {}): Promise<string> {
	const { env } = await import("../env");
	const { dir: tmpDir } = await import("tmp");
	const { dir = env.temp, prefix, template } = options;
	const tmpOptions: {
		dir?: string;
		tmpdir?: string;
		prefix?: string;
		template?: string;
	} = { prefix, template };

	if (isAbsolute(dir)) {
		tmpOptions.tmpdir = dir;
	} else {
		tmpOptions.dir = dir;
	}

	await ensureDir(dir);
	return new Promise<string>((resolve, reject) => {
		tmpDir(tmpOptions, (err: any, path: string) => {
			if (err) return reject(err);
			resolve(path);
		});
	});
}

// (for mocking only)
function reset() {}

export { readFileSync, watch } from "fs";
export {
	checksum,
	copy,
	emptyDir,
	ensureDir,
	ensureDirSync,
	move,
	pathExists,
	readFile,
	readJson,
	remove,
	reset,
	symlink,
	tmpFile,
	tmpFolder,
	writeFile
};
