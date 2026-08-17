/**
 * Jest globalSetup for the e2e suite (visible mode only).
 *
 * Ensures the freshly-built `vbapm.xlam` add-in is open in a visible Excel
 * instance before tests run, so the `run.ps1` bridge can attach to it (and reuse
 * it) rather than opening/closing it per macro invocation. Also verifies the
 * open add-in points at the repo build (not a stale %APPDATA% copy).
 *
 * This runs in a separate Node process (Jest's globalSetup), so any COM
 * side-effects (opening Excel + the add-in) persist for the test process, which
 * attaches via GetActiveObject.
 *
 * In background mode (VBA_BACKGROUND_BUILD=1) this is a no-op: each test spin-up
 * its own hidden instance and we don't reuse a visible one.
 */
const { spawnSync } = require("child_process");
const { join } = require("path");
const fs = require("fs");

module.exports = async function globalSetup() {
	writeRunBanner();

	const background = /^(1|true|yes)$/i.test(process.env.VBA_BACKGROUND_BUILD || "");

	if (background) {
		// Nothing to pre-open — background mode uses per-test hidden instances.
		return;
	}

	const script = join(__dirname, "..", "scripts", "ps", "EnsureVbapmAddin.ps1");

	const result = spawnSync(
		"powershell.exe",
		["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script],
		{ stdio: "inherit" }
	);

	if (result.status !== 0) {
		// Best-effort: don't fail the whole suite just because the pre-open
		// couldn't run; print a warning and continue.
		console.warn("[e2e globalSetup] EnsureVbapmAddin.ps1 exited non-zero; continuing anyway.");
	}
};

/**
 * Write a clear "run banner" separator into the instance log so each e2e run is
 * visually distinct when reviewing `instances.log`. This is done in globalSetup
 * (once per suite, before tests) because `run.ps1` is a short-lived process per
 * macro call and cannot detect run boundaries on its own.
 */
function writeRunBanner() {
	try {
		const logPath =
			process.env.VBA_INSTANCE_LOG || join(process.env.TEMP, "Excel-Instances", "instances.log");

		fs.mkdirSync(require("path").dirname(logPath), { recursive: true });

		const mode = /^(1|true|yes)$/i.test(process.env.VBA_BACKGROUND_BUILD || "")
			? "background"
			: "visible";
		const inProcess = /^(1|true|yes)$/i.test(process.env.E2E_IN_PROCESS || "")
			? " in-process"
			: " spawned";

		const stamp = new Date().toISOString();
		const banner = [
			"",
			`${"=".repeat(70)}`,
			`=== RUN START  ${stamp}  mode=${mode}${inProcess}`,
			`${"=".repeat(70)}`
		].join("\n");

		fs.appendFileSync(logPath, banner + "\n");
	} catch {
		// Banner is best-effort; never fail the suite over logging.
	}
}
