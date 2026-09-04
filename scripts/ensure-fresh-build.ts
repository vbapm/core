/**
 * Checks whether the generated library and add-in are up to date.
 *
 * This entry point runs freshness checks in dependency order:
 *
 * 1. Rebuilds `lib/` when the TypeScript sources or build configuration are newer.
 * 2. Checks `addins/build/vbapm.xlam` when its VBA or XML sources are newer.
 * 3. With `--force`, checks `scripts/bootstrap/build/bootstrap.xlsm` when the
 *    add-in or its related sources are newer.
 *
 * The add-in check watches only files under `addins/`, so TypeScript changes do
 * not trigger an Office-backed rebuild by default. The `--force` flag forces
 * the add-in check and includes the bootstrap check. It does not bypass the
 * bootstrap freshness record. If a check fails, this script exits with the
 * child's exit code and does not run the remaining checks.
 *
 * @remarks
 * The checks run as separate Node processes so each one can use its own
 * freshness metadata and build command while this script provides one entry
 * point for commands that need fresh build artifacts. Use `--force` when a
 * TypeScript change is known to affect the generated Office artifacts.
 */
const { execFileSync } = require("node:child_process");
const { join, resolve } = require("node:path");

const ROOT = resolve(__dirname, "..");
const CHECKS = ["ensure-fresh-lib.ts", "ensure-fresh-addin.ts"];
const FORCE_CHECKS = ["ensure-fresh-bootstrap.ts"];
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
	const force = process.argv.includes("--force");
	const checks = force ? [...CHECKS, ...FORCE_CHECKS] : CHECKS;

	for (const check of checks) {
		console.log(`[ensure-fresh-build] ${CHECK_EMOJIS[check]} running ${check}...`);
		try {
			const args =
				force && check === "ensure-fresh-addin.ts"
					? [join(__dirname, check), "--force"]
					: [join(__dirname, check)];
			execFileSync(process.execPath, args, {
				cwd: ROOT,
				stdio: "inherit"
			});
		} catch (err: unknown) {
			process.exit(exitCode(err));
		}
	}

	console.log("[ensure-fresh-build] all requested build artifacts are fresh.");
}

main();

type CommonJsEntryPoint = never;
export type { CommonJsEntryPoint };
