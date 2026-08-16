/**
 * Cross-platform guard for the e2e integration suite.
 *
 * On Windows, verifies the Excel instance coordination registry is clean before
 * allowing the suite to run — it fails fast if any EXCEL.EXE process is running
 * but NOT tracked in the registry (a rogue/user/other-agent session). This
 * prevents the e2e suite (which drives Excel via COM) from hijacking or
 * interfering with those sessions.
 *
 * On macOS and other platforms, this is a no-op: the Excel instance
 * coordination registry is a Windows-only mechanism (macOS uses AppleScript via
 * run.applescript and does not share a per-machine COM instance pool in the
 * same way).
 *
 * Run natively by Node 24 type-stripping (like scripts/ensure-fresh-build.ts).
 * Must use CommonJS `require` + `__dirname` (the repo package.json has no
 * "type": "module").
 */
const { spawnSync } = require("child_process");
const { join } = require("path");

function main() {
	if (process.platform !== "win32") {
		// Non-Windows: nothing to coordinate, nothing to check.
		process.exit(0);
	}

	const statusScript = join(__dirname, "ps", "Ensure-ExcelInstancesClean.ps1");
	const result = spawnSync(
		"powershell.exe",
		["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", statusScript],
		{ stdio: "inherit" }
	);

	if (result.error) {
		console.error(`Failed to launch Excel instance check: ${result.error.message}`);
		process.exit(1);
	}

	if (result.status !== 0) {
		console.error(
			"Excel instance registry is not clean. Rogue Excel instances detected — " +
				"aborting e2e suite to avoid interfering with another session. " +
				"Review %TEMP%\\Excel-Instances\\instances.json and close/register stray EXCEL.EXE processes, then retry."
		);
		process.exit(result.status ?? 1);
	}

	process.exit(0);
}

main();
