/**
 * Jest globalTeardown for the e2e suite.
 *
 * Runs the end-of-suite Excel instance assessment: reconcile the coordination
 * registry, then report any instances that linger longer than expected or are
 * "zombies" (deactivated in the registry but whose EXCEL.EXE process is still
 * alive). The report identifies the workbooks each flagged instance held, so a
 * zombie can be traced back to the e2e test that opened them.
 *
 * Report-only by default so a race-y Excel Quit() doesn't flake CI. Set
 * VBA_FAIL_ON_ZOMBIE=1 to fail the suite when zombies are detected.
 */
const { spawnSync } = require("child_process");
const { join } = require("path");

module.exports = async function globalTeardown() {
	const script = join(__dirname, "..", "scripts", "ps", "Assess-ExcelInstances.ps1");

	const args = ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script];
	if (/^(1|true|yes)$/i.test(process.env.VBA_FAIL_ON_ZOMBIE || "")) {
		args.push("-FailOnZombie");
	}

	const result = spawnSync("powershell.exe", args, { stdio: "inherit" });

	if (result.status !== 0) {
		// The assessment script only exits non-zero when -FailOnZombie is set and
		// zombies were found. Surface that as a suite failure.
		throw new Error(
			`Assess-ExcelInstances.ps1 exited with status ${result.status} (zombie Excel instances detected).`
		);
	}
};
