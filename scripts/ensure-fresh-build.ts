/**
 * Checks whether the generated library and add-in artifacts are up to date.
 *
 * This entry point runs freshness checks in dependency order:
 *
 * 1. Rebuilds `lib/` when the TypeScript sources or build configuration are newer.
 * 2. Rebuilds `addins/build/vbapm.xlam` when its VBA sources are newer.
 * 3. With `--force`, rebuilds `scripts/bootstrap/build/bootstrap.xlsm` when the
 *    add-in or its related sources are newer.
 *
 * Each child script skips its rebuild when its output is current. The
 * `--force` flag includes the expensive bootstrap check; it does not bypass
 * that check's freshness record. If a check fails, this script exits with the
 * child's exit code and does not run the remaining checks.
 *
 * @remarks
 * The checks run as separate Node processes so each one can use its own
 * freshness metadata and build command while this script provides one entry
 * point for commands that need fresh build artifacts. The default omits the
 * bootstrap workbook for faster local builds.
 */
const { execFileSync } = require("node:child_process");
const { join, resolve } = require("node:path");

const ROOT = resolve(__dirname, "..");
const CHECKS = ["ensure-fresh-lib.ts", "ensure-fresh-addin.ts"];
const BOOTSTRAP_CHECK = "ensure-fresh-bootstrap.ts";
const CHECK_EMOJIS: Record<string, string> = {
	"ensure-fresh-lib.ts": "📦",
	"ensure-fresh-addin.ts": "📎",
	"ensure-fresh-bootstrap.ts": "🥾"
};

function exitCode(err: unknown): number {
	if (err && typeof err === "object" && "status" in err) {
		const status = (err as { status: unknown }).status;
		if (typeof status === "number") return status;
	}
	return 1;
}

function main(): void {
	const checks = process.argv.includes("--force") ? [...CHECKS, BOOTSTRAP_CHECK] : CHECKS;

	for (const check of checks) {
		console.log(`[ensure-fresh-build] ${CHECK_EMOJIS[check]} running ${check}...`);
		try {
			execFileSync(process.execPath, [join(__dirname, check)], {
				cwd: ROOT,
				stdio: "inherit"
			});
		} catch (err: unknown) {
			process.exit(exitCode(err));
		}
	}

	console.log("[ensure-fresh-build] all build outputs are fresh.");
}

main();

type CommonJsEntryPoint = never;
export type { CommonJsEntryPoint };
