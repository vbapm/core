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
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = fileURLToPath(new URL(".", import.meta.url));

function teardownTrace(message) {
	if (!/^(1|true|yes)$/i.test(process.env.VBA_SESSION_TRACE || "")) return;
	try {
		const path = join(process.env.TEMP, "Excel-Instances", "session-client-trace.log");
		fs.appendFileSync(path, `${new Date().toISOString()} global-teardown ${message}\n`);
	} catch {}
}

export default async function globalTeardown() {
	teardownTrace("start");
	const script = join(__dirname, "..", "scripts", "ps", "Assess-ExcelInstances.ps1");

	const args = ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script];
	if (/^(1|true|yes)$/i.test(process.env.VBA_FAIL_ON_ZOMBIE || "")) {
		args.push("-FailOnZombie");
	}

	const result = spawnSync("powershell.exe", args, { stdio: "inherit" });
	teardownTrace(`assessment-complete status=${result.status}`);

	if (result.status !== 0) {
		// The assessment script only exits non-zero when -FailOnZombie is set and
		// zombies were found. Surface that as a suite failure.
		throw new Error(
			`Assess-ExcelInstances.ps1 exited with status ${result.status} (zombie Excel instances detected).`
		);
	}

	await cleanupTrackedE2EInstances();
	teardownTrace("tracked-cleanup-complete");
	await sweepTemporaryRepos();
	teardownTrace("temporary-sweep-complete");
	finishEvolutionRun();
	teardownTrace("run-info-complete");
};

/**
 * Stamp the finish time + total elapsed duration onto the evolution run's
 * `run-info.json` (created by tests/__helpers__/execute.ts at run start). The
 * marker `current-run.txt` points at the run directory that received the
 * snapshots, so this works even though global-teardown runs without
 * `VBA_EVOLUTION_RUN_ID` set.
 */
function finishEvolutionRun() {
	if (!/^(1|true|yes)$/i.test(process.env.VBA_EVOLUTION || "")) return;
	try {
		const marker = join(process.env.TEMP, "Excel-Instances", "evolution", "current-run.txt");
		if (!fs.existsSync(marker)) return;
		const runDir = fs.readFileSync(marker, "utf8").trim();
		if (!runDir || !fs.existsSync(runDir)) return;

		const infoPath = join(runDir, "run-info.json");
		if (!fs.existsSync(infoPath)) return;
		const info = JSON.parse(fs.readFileSync(infoPath, "utf8"));
		info.finishedAt = new Date().toISOString();
		const start = new Date(info.startedAt).getTime();
		if (Number.isFinite(start)) {
			info.elapsedMs = Date.now() - start;
		}
		fs.writeFileSync(infoPath, JSON.stringify(info, null, 2), "utf8");
		teardownTrace(`run-info finished elapsedMs=${info.elapsedMs}`);
	} catch (err) {
		teardownTrace(`run-info finish failed: ${err.message}`);
	}
}

async function cleanupTrackedE2EInstances() {
	const script = join(__dirname, "..", "scripts", "ps", "Close-TrackedE2EExcelInstances.ps1");
	const persistentSessionValue = process.env.VBA_PERSISTENT_SESSION || "";
	const persistentSessionLimit = /^(true|yes)$/i.test(persistentSessionValue)
		? 1
		: Number.parseInt(persistentSessionValue, 10);
	const keepWarm =
		/^(1|true|yes)$/i.test(process.env.E2E_BACKGROUND || "") &&
		Number.isInteger(persistentSessionLimit) &&
		persistentSessionLimit > 0;
	const args = ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script];
	if (keepWarm) args.push("-KeepWarmCount", String(persistentSessionLimit));
	const result = spawnSync("powershell.exe", args, {
		stdio: "inherit"
	});

	if (result.status !== 0) {
		console.warn(`[e2e] tracked Excel cleanup exited with status ${result.status}`);
	}
}

function registryPath() {
	return join(process.env.TEMP || ".", "Excel-Instances", "instances.json");
}

function normalizePathJs(value) {
	return String(value).replace(/\\/g, "/").toLowerCase();
}

/** File paths (workbooks + loaded add-ins) every ACTIVE keep-warm instance holds. */
function warmHeldPathSet() {
	try {
		const doc = JSON.parse(fs.readFileSync(registryPath(), "utf8"));
		const held = [];
		for (const inst of Array.isArray(doc.active) ? doc.active : []) {
			const warm = inst.intent === "keep-warm" || String(inst.owner || "").startsWith("session#");
			if (!warm) continue;
			const workbooks = inst.workbooks;
			if (Array.isArray(workbooks)) {
				for (const wb of workbooks) {
					if (wb) held.push(normalizePathJs(wb));
				}
			} else if (typeof workbooks === "string" && workbooks) {
				held.push(normalizePathJs(workbooks));
			}
			if (Array.isArray(inst.addins)) {
				for (const addin of inst.addins) {
					if (addin && addin.fullPath) held.push(normalizePathJs(addin.fullPath));
				}
			}
		}
		return new Set(held);
	} catch {
		return new Set();
	}
}

function isWarmHeld(path, warmHeld) {
	const prefix = normalizePathJs(path).replace(/\/+$/, "") + "/";
	for (const held of warmHeld) {
		if (held.startsWith(prefix)) return true;
	}
	return false;
}

async function sweepTemporaryRepos() {
	if (/^(1|true|yes)$/i.test(process.env.KEEP_E2E_TMP || "")) {
		return;
	}

	const tempRoot = join(__dirname, ".tmp");
	if (!fs.existsSync(tempRoot)) {
		return;
	}

	const started = Date.now();
	// A keep-warm Excel instance may still be alive (teardown deliberately keeps
	// one warm) holding a file inside a leftover temp folder. Deleting such a
	// folder is futile while that instance lives — skip the attempt and leave it
	// for the `clean` script (which closes invisible Excel instances first).
	const warmHeld = warmHeldPathSet();
	teardownTrace(`sweep start root=${tempRoot} warmHeld=${warmHeld.size}`);
	for (const entry of fs.readdirSync(tempRoot)) {
		const path = join(tempRoot, entry);
		if (isWarmHeld(path, warmHeld)) {
			console.warn(
				`[e2e] temporary repo sweep skipped (held by keep-warm Excel instance): ${path}`
			);
			teardownTrace(`sweep skipped-held-warm path=${path}`);
			continue;
		}
		try {
			teardownTrace(`sweep attempt path=${path}`);
			await fs.promises.rm(path, {
				recursive: true,
				force: true,
				maxRetries: 1,
				retryDelay: 250
			});
			teardownTrace(`sweep removed path=${path}`);
		} catch (err) {
			console.warn(`[e2e] final temporary repo sweep deferred: ${path} (${err.message})`);
			teardownTrace(`sweep deferred path=${path} (${err.message})`);
		}
	}
	teardownTrace(`sweep complete elapsed=${Date.now() - started}ms`);
}
