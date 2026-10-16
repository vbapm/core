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
 * In background mode (E2E_BACKGROUND=1) this is a no-op: each test spins up
 * its own hidden instance and we don't reuse a visible one.
 */
import { spawnSync } from "node:child_process";
import { join, dirname } from "node:path";
import fs from "node:fs";
import { fileURLToPath } from "node:url";

const __dirname = fileURLToPath(new URL(".", import.meta.url));

export default async function globalSetup() {
	recordEvolutionRunStart();
	clearActiveE2EInstanceLogs();
	ensureExcelInstancesClean();
	clearInactiveInstances();
	writeRunBanner();

	const background = /^(1|true|yes)$/i.test(process.env.E2E_BACKGROUND || "");

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
 * Write a `run-info.json` (configuration flags + start time) into the current
 * evolution run directory — the one chosen by Clear-ExcelRegistryEvolution.ps1
 * and recorded in `current-run.txt`. The Excel registry writes all of the run's
 * evolution snapshots to that same directory (no per-worker RUN_ID override),
 * and `global-teardown` later stamps the finish time + elapsed duration so the
 * HTML report header can show how the suite was configured and how long it ran.
 */
function recordEvolutionRunStart() {
	if (!/^(1|true|yes)$/i.test(process.env.VBA_EVOLUTION || "")) return;
	try {
		const evolutionDir = join(process.env.TEMP, "Excel-Instances", "evolution");
		const marker = join(evolutionDir, "current-run.txt");
		if (!fs.existsSync(marker)) return;
		const runDir = fs.readFileSync(marker, "utf8").trim();
		if (!runDir) return;
		fs.mkdirSync(runDir, { recursive: true });

		const flags = {};
		for (const [key, value] of Object.entries(process.env)) {
			if (!value) continue;
			if (/^(E2E_|VBA_)/.test(key) || key === "KEEP_E2E_TMP") flags[key] = value;
		}

		const info = {
			runId: runDir.split(/[\\/]/).pop(),
			startedAt: new Date().toISOString(),
			script: process.env.npm_lifecycle_event || "",
			command: process.env.npm_lifecycle_script || process.argv.slice(2).join(" "),
			flags
		};
		fs.writeFileSync(join(runDir, "run-info.json"), JSON.stringify(info, null, 2), "utf8");
	} catch (err) {
		console.warn(`[e2e globalSetup] failed to record evolution run start: ${err.message}`);
	}
}

function ensureExcelInstancesClean() {
	if (process.platform !== "win32") return;

	const script = join(__dirname, "..", "scripts", "ps", "Ensure-ExcelInstancesClean.ps1");
	const clean = /^(1|true|yes)$/i.test(process.env.E2E_CLEAN || "");
	const args = ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script];
	if (clean) args.push("-Clean");
	const startedAt = new Date().toISOString();
	const result = spawnSync("powershell.exe", args, { encoding: "utf8" });
	if (result.stdout) process.stdout.write(result.stdout);
	if (result.stderr) process.stderr.write(result.stderr);
	recordEnsureExcelInstancesClean({
		startedAt,
		completedAt: new Date().toISOString(),
		status: result.status,
		logs: [result.stdout, result.stderr, result.error?.message].filter(Boolean).flatMap(output =>
			output
				.split(/\r?\n/)
				.map(line => line.trim())
				.filter(Boolean)
		)
	});

	if (result.status !== 0) {
		throw new Error("Rogue Excel instances remain alive; refusing to start e2e tests.");
	}
}

function recordEnsureExcelInstancesClean(event) {
	if (!/^(1|true|yes)$/i.test(process.env.VBA_EVOLUTION || "")) return;
	try {
		const evolutionDir = join(process.env.TEMP, "Excel-Instances", "evolution");
		const marker = join(evolutionDir, "current-run.txt");
		if (!fs.existsSync(marker)) return;
		const runDir = fs.readFileSync(marker, "utf8").trim();
		if (!runDir) return;

		const infoPath = join(runDir, "run-info.json");
		const info = JSON.parse(fs.readFileSync(infoPath, "utf8"));
		info.ensureExcelInstancesClean = event;
		fs.writeFileSync(infoPath, JSON.stringify(info, null, 2), "utf8");
	} catch (err) {
		console.warn(`[e2e globalSetup] failed to record preflight details: ${err.message}`);
	}
}

/**
 * Reset the registry's inactive (recently deactivated) list before tests run so
 * the end-of-suite assessment only reports instances deactivated during this run.
 */
function clearInactiveInstances() {
	const script = join(__dirname, "..", "scripts", "ps", "Clear-InactiveExcelInstances.ps1");

	const result = spawnSync(
		"powershell.exe",
		["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script],
		{ stdio: "ignore" }
	);

	if (result.status !== 0) {
		// Best-effort: don't fail the suite over registry housekeeping.
		console.warn(
			"[e2e globalSetup] Clear-InactiveExcelInstances.ps1 exited non-zero; continuing anyway."
		);
	}
}

function clearActiveE2EInstanceLogs() {
	const script = join(__dirname, "..", "scripts", "ps", "Clear-ActiveE2EExcelInstanceLogs.ps1");
	const result = spawnSync(
		"powershell.exe",
		["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script],
		{ stdio: "ignore" }
	);

	if (result.status !== 0) {
		console.warn(
			"[e2e globalSetup] Clear-ActiveE2EExcelInstanceLogs.ps1 exited non-zero; continuing anyway."
		);
	}
}

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

		fs.mkdirSync(dirname(logPath), { recursive: true });

		const mode = /^(1|true|yes)$/i.test(process.env.E2E_BACKGROUND || "")
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
