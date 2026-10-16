/**
 * Ensures the compiled `lib/` folder is up to date before running tests or
 * deploying the development build.
 *
 * This script compares the modification time of the build outputs against
 * every source input. If any input is newer than the output, the build is
 * considered stale and is regenerated.
 *
 * Outputs (must all exist and be newer than every input):
 *   - `lib/vbapm.js` (CLI entry point, used by `bin/vba`)
 *   - `lib/index.js` (library entry point, mapped by e2e Jest config)
 *
 * Inputs (any of these newer than an output marks the build stale):
 *   - all files under `src/` (code + templates + submodules)
 *   - `rollup.config.mjs` (build configuration)
 *   - `package.json` (`dependencies` changes can affect the bundle)
 *
 * Without `--dev`, it checks for a production build and runs `pnpm run build:cli`
 * when stale. With `--dev`, it checks for a development build and runs
 * `pnpm run build:dev` when stale. A mode record prevents either build from
 * being mistaken for the other.
 */
import { execSync } from "node:child_process";
import { readFileSync, readdirSync, statSync, existsSync, writeFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(fileURLToPath(new URL("..", import.meta.url)));
const BUILD_EMOJI = "📦";
const BUILD_MODE = process.argv.includes("--dev") ? "development" : "production";
const BUILD_MODE_FILE = join(ROOT, "lib", ".build-mode");
const BUILD_COMMAND = BUILD_MODE === "development" ? "pnpm run build:dev" : "pnpm run build:cli";

const OUTPUTS = ["lib/vbapm.js", "lib/index.js"];
const INPUT_ROOTS = ["src"];
const EXTRA_INPUTS = ["rollup.config.mjs", "package.json"];

function walk(dir: string, acc: string[] = []): string[] {
	let entries: ReturnType<typeof readdirSync>;
	try {
		entries = readdirSync(dir, { withFileTypes: true });
	} catch {
		return acc;
	}

	for (const entry of entries) {
		const full = join(dir, entry.name);
		if (entry.isDirectory()) {
			walk(full, acc);
		} else if (entry.isFile()) {
			acc.push(full);
		}
	}

	return acc;
}

function latestMtime(paths: string[]): number {
	let latest = 0;
	for (const p of paths) {
		if (!existsSync(p)) continue;
		try {
			const mtime = statSync(p).mtimeMs;
			if (mtime > latest) latest = mtime;
		} catch {
			// Ignore unreadable files.
		}
	}
	return latest;
}

function earliestMtime(paths: string[]): number | null {
	let earliest: number | null = null;
	for (const p of paths) {
		if (!existsSync(p)) return null;
		try {
			const mtime = statSync(p).mtimeMs;
			if (earliest === null || mtime < earliest) earliest = mtime;
		} catch {
			return null;
		}
	}
	return earliest;
}

function isStale(): { stale: boolean; reason?: string } {
	const outputs = OUTPUTS.map(p => join(ROOT, p));
	const outputMtime = earliestMtime(outputs);
	if (outputMtime === null) {
		const missing = OUTPUTS.filter(p => !existsSync(join(ROOT, p)));
		return { stale: true, reason: `missing build output: ${missing.join(", ")}` };
	}

	if (!existsSync(BUILD_MODE_FILE)) {
		return { stale: true, reason: `missing build mode record: ${BUILD_MODE_FILE}` };
	}

	if (readFileSync(BUILD_MODE_FILE, "utf8").trim() !== BUILD_MODE) {
		return { stale: true, reason: `build output is not a ${BUILD_MODE} build` };
	}

	const inputs = INPUT_ROOTS.flatMap(r => walk(join(ROOT, r)));
	for (const p of EXTRA_INPUTS) {
		inputs.push(join(ROOT, p));
	}

	const inputMtime = latestMtime(inputs);
	if (inputMtime > outputMtime) {
		const staleInputs = inputs
			.filter(p => existsSync(p) && statSync(p).mtimeMs > outputMtime)
			.slice(0, 5);
		return {
			stale: true,
			reason: `sources newer than build output (e.g. ${staleInputs
				.map(p => p.replace(ROOT + "/", "").replace(ROOT + "\\", ""))
				.join(", ")})`
		};
	}

	return { stale: false };
}

function main(): void {
	const { stale, reason } = isStale();

	if (!stale) {
		console.log(`[ensure-fresh-lib] ${BUILD_EMOJI} lib/ is up to date - skipping rebuild 🚀`);
		return;
	}

	console.log(`[ensure-fresh-lib] ${BUILD_EMOJI} lib/ is stale (${reason}) - rebuilding...`);
	try {
		execSync(BUILD_COMMAND, { cwd: ROOT, stdio: "inherit" });
	} catch (err: unknown) {
		const code =
			err && typeof err === "object" && "status" in err ? (err as { status: number }).status : 1;
		process.exit(typeof code === "number" ? code : 1);
	}

	writeFileSync(BUILD_MODE_FILE, `${BUILD_MODE}\n`, "utf8");

	console.log(`[ensure-fresh-lib] ${BUILD_EMOJI} rebuild complete.`);
}

main();
