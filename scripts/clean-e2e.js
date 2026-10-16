/**
 * Full e2e cleanup: close every invisible Excel instance, then delete the
 * temporary e2e repos under tests/.tmp.
 *
 * Temp folders whose workbooks were opened by a keep-warm Excel instance cannot
 * be deleted while that instance lives (Excel keeps the file/folder locked), so
 * the normal teardown sweep deliberately leaves them behind. This script is the
 * "future cleanup" that removes them once no invisible Excel holds them.
 *
 * Visible Excel sessions (a user's own workbooks) are never touched.
 */
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = fileURLToPath(new URL("..", import.meta.url));
const tmpRoot = path.join(repoRoot, "tests", ".tmp");

function closeInvisibleExcel() {
	const script = path.join(repoRoot, "scripts", "ps", "Close-AllInvisibleExcelInstances.ps1");
	const result = spawnSync(
		"powershell.exe",
		["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script],
		{ stdio: "inherit", cwd: repoRoot }
	);
	if (result.status !== 0) {
		console.warn(
			`[clean] Close-AllInvisibleExcelInstances.ps1 exited with status ${result.status}`
		);
	}
}

function sleep(ms) {
	return new Promise(resolve => setTimeout(resolve, ms));
}

async function main() {
	console.log("[clean] closing invisible Excel instances...");
	closeInvisibleExcel();

	// Give Excel a moment to release its file handles before removing folders.
	await sleep(1500);

	console.log(`[clean] removing temporary e2e repos under ${tmpRoot}...`);
	if (!fs.existsSync(tmpRoot)) {
		console.log("[clean] nothing to clean (tests/.tmp does not exist).");
		return;
	}

	let removed = 0;
	let deferred = 0;
	for (const entry of fs.readdirSync(tmpRoot)) {
		const target = path.join(tmpRoot, entry);
		try {
			fs.rmSync(target, { recursive: true, force: true, maxRetries: 2, retryDelay: 250 });
			removed++;
		} catch (err) {
			deferred++;
			console.warn(`[clean] could not remove ${target}: ${err.message}`);
		}
	}
	console.log(
		`[clean] done: removed ${removed} item(s)${deferred ? `, deferred ${deferred}` : ""}.`
	);
}

main().catch(err => {
	console.error(`[clean] failed: ${err.message}`);
	process.exitCode = 1;
});
