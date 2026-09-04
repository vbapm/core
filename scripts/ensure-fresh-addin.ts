const { execSync } = require("node:child_process");
const { existsSync } = require("node:fs");
const { join } = require("node:path");
const { ROOT, checkFreshness, displayPath, writeFreshness } = require("./freshness.ts");

const BUILD_EMOJI = "📎";
const OUTPUT = join(ROOT, "addins", "build", "vbapm.xlam");
const METADATA = `${OUTPUT}.fresh.json`;
const BOOTSTRAP = join(ROOT, "scripts", "bootstrap", "build", "bootstrap.xlsm");

function exitCode(err: unknown): number {
	if (err && typeof err === "object" && "status" in err) {
		const status = (err as { status: unknown }).status;
		if (typeof status === "number") return status;
	}
	return 1;
}

function main(): void {
	if (!existsSync(BOOTSTRAP)) {
		console.error(
			`[ensure-fresh-addin] ${BUILD_EMOJI} missing bootstrap host: ${displayPath(BOOTSTRAP)}`
		);
		process.exit(1);
	}

	const result = checkFreshness(OUTPUT, METADATA);
	if (result.fresh) {
		console.log(
			`[ensure-fresh-addin] ${BUILD_EMOJI} ${displayPath(OUTPUT)} is up to date - skipping rebuild 🚀`
		);
		return;
	}

	console.log(
		`[ensure-fresh-addin] ${BUILD_EMOJI} ${displayPath(OUTPUT)} is stale (${result.reason}) - rebuilding...`
	);
	try {
		execSync("pnpm run build:addins", { cwd: ROOT, stdio: "inherit" });
	} catch (err: unknown) {
		process.exit(exitCode(err));
	}

	if (!existsSync(OUTPUT)) {
		console.error(
			`[ensure-fresh-addin] ${BUILD_EMOJI} build completed without creating ${displayPath(OUTPUT)}.`
		);
		process.exit(1);
	}

	writeFreshness(METADATA);
	console.log(`[ensure-fresh-addin] ${BUILD_EMOJI} ${displayPath(OUTPUT)} is fresh.`);
}

main();

type CommonJsEntryPoint = never;
export type { CommonJsEntryPoint };
