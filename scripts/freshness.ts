const { createHash } = require("node:crypto");
const { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } = require("node:fs");
const { dirname, join, relative, resolve } = require("node:path");

const ROOT = resolve(__dirname, "..");
const SCHEMA_VERSION = 1;
const ALGORITHM = "sha256";

const INPUT_ROOTS = ["addins", "scripts/bootstrap", "lib"];
const ADDIN_INPUT_ROOTS = ["addins"];
const EXTRA_INPUTS = ["package.json", "pnpm-lock.yaml", "rollup.config.mjs"];
const SCRIPT_INPUTS = [
	"scripts/build-addins.js",
	"scripts/ensure-fresh-addin.ts",
	"scripts/ensure-fresh-bootstrap.ts",
	"scripts/ensure-fresh-build.ts",
	"scripts/ensure-fresh-lib.ts",
	"scripts/freshness.ts"
];
const REQUIRED_LIB_OUTPUTS = ["lib/vbapm.js", "lib/index.js"];
const EXCLUDED_ROOTS = ["addins/build", "scripts/bootstrap/build"];
type FingerprintScope = "all" | "addin";

function normalizeRepoPath(file: string): string {
	return relative(ROOT, resolve(file)).replace(/\\/g, "/");
}

function isExcluded(file: string): boolean {
	const path = normalizeRepoPath(file);
	return EXCLUDED_ROOTS.some(root => path === root || path.startsWith(`${root}/`));
}

function walk(dir: string, files: string[] = []): string[] {
	if (!existsSync(dir)) return files;

	let entries: Array<{ name: string; isDirectory: () => boolean; isFile: () => boolean }>;
	try {
		entries = readdirSync(dir, { withFileTypes: true }) as typeof entries;
		entries.sort((left, right) => (left.name < right.name ? -1 : left.name > right.name ? 1 : 0));
	} catch {
		return files;
	}

	for (const entry of entries) {
		const full = join(dir, entry.name);
		if (entry.isDirectory()) {
			if (!isExcluded(full)) walk(full, files);
		} else if (entry.isFile() && !isExcluded(full)) {
			files.push(full);
		}
	}

	return files;
}

function getFingerprintInputs(scope: FingerprintScope = "all"): string[] {
	const paths: string[] = [];
	const inputRoots = scope === "addin" ? ADDIN_INPUT_ROOTS : INPUT_ROOTS;
	const extraInputs = scope === "addin" ? [] : EXTRA_INPUTS;
	const scriptInputs = scope === "addin" ? [] : SCRIPT_INPUTS;
	const requiredLibOutputs = scope === "addin" ? [] : REQUIRED_LIB_OUTPUTS;

	for (const root of inputRoots) {
		walk(join(ROOT, root), paths);
	}
	for (const input of extraInputs) {
		paths.push(join(ROOT, input));
	}
	for (const input of scriptInputs) {
		paths.push(join(ROOT, input));
	}
	for (const output of requiredLibOutputs) {
		paths.push(join(ROOT, output));
	}

	return sortPaths(paths);
}

function sortPaths(paths: string[]): string[] {
	const unique = [...new Set(paths)];
	return unique.sort((left, right) => {
		const leftPath = normalizeRepoPath(left);
		const rightPath = normalizeRepoPath(right);
		return leftPath < rightPath ? -1 : leftPath > rightPath ? 1 : 0;
	});
}

function computeFingerprint(inputPaths: string[] = getFingerprintInputs()): string {
	const hash = createHash(ALGORITHM);

	for (const input of sortPaths(inputPaths)) {
		const full = resolve(input);
		const path = normalizeRepoPath(full);
		hash.update(`path:${path}\0`);

		if (!existsSync(full)) {
			hash.update("missing\0");
			continue;
		}

		const data = readFileSync(full);
		hash.update(`length:${data.length}\0`);
		hash.update(data);
		hash.update("\0");
	}

	return `${ALGORITHM}-${hash.digest("hex")}`;
}

function createFreshnessRecord(scope: FingerprintScope = "all") {
	const inputPaths = getFingerprintInputs(scope);

	return {
		schemaVersion: SCHEMA_VERSION,
		algorithm: ALGORITHM,
		fingerprint: computeFingerprint(inputPaths),
		inputs: inputPaths.map(normalizeRepoPath)
	};
}

function displayPath(file: string): string {
	return normalizeRepoPath(file);
}

function readFreshnessRecord(file: string) {
	try {
		return JSON.parse(readFileSync(file, "utf8"));
	} catch {
		return null;
	}
}

function checkFreshness(
	output: string,
	metadata: string,
	scope: FingerprintScope = "all"
): { fresh: boolean; reason: string } {
	if (!existsSync(output)) {
		return { fresh: false, reason: `missing build output: ${displayPath(output)}` };
	}
	if (!existsSync(metadata)) {
		return { fresh: false, reason: `missing freshness record: ${displayPath(metadata)}` };
	}

	const record = readFreshnessRecord(metadata);
	if (
		!record ||
		record.schemaVersion !== SCHEMA_VERSION ||
		record.algorithm !== ALGORITHM ||
		typeof record.fingerprint !== "string" ||
		!Array.isArray(record.inputs)
	) {
		return { fresh: false, reason: `invalid freshness record: ${displayPath(metadata)}` };
	}

	const current = createFreshnessRecord(scope);
	if (
		record.fingerprint !== current.fingerprint ||
		record.inputs.join("\n") !== current.inputs.join("\n")
	) {
		return { fresh: false, reason: `source fingerprint changed for ${displayPath(output)}` };
	}

	return { fresh: true, reason: "" };
}

function writeFreshness(metadata: string, scope: FingerprintScope = "all") {
	const record = createFreshnessRecord(scope);
	mkdirSync(dirname(metadata), { recursive: true });
	writeFileSync(metadata, `${JSON.stringify(record, null, 2)}\n`, "utf8");
	return record;
}

module.exports = {
	ROOT,
	checkFreshness,
	computeFingerprint,
	createFreshnessRecord,
	displayPath,
	getFingerprintInputs,
	writeFreshness
};
