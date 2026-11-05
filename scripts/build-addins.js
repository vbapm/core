import { execFileSync } from "node:child_process";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import * as vba from "../lib/index.js";

const __dirname = fileURLToPath(new URL(".", import.meta.url));

main().catch(err => {
	console.error(err.message);
	if (err.underlying && err.underlying.result) {
		console.log(err.underlying.result);
	}
	process.exit(1);
});

async function main() {
	closeOpenAddin();

	// Hidden instance: a visible Excel may have the installed vbapm.xlam loaded under the same name
	await vba.buildProject({
		addin: join(__dirname, "bootstrap/build/bootstrap.xlsm"),
		background: true
	});
}

function closeOpenAddin() {
	const addinPath = join(__dirname, "..", "addins", "build", "vbapm.xlam");

	if (process.platform === "win32") {
		execFileSync(
			"powershell.exe",
			[
				"-NoProfile",
				"-ExecutionPolicy",
				"Bypass",
				"-File",
				join(__dirname, "ps", "Close-VbapmAddin.ps1"),
				"-AddinPath",
				addinPath
			],
			{ stdio: "inherit" }
		);
	} else if (process.platform === "darwin") {
		execFileSync(
			"osascript",
			[
				"-e",
				[
					"on run argv",
					"set addinPath to item 1 of argv",
					'tell application "System Events" to if not (exists process "Microsoft Excel") then return',
					'tell application "Microsoft Excel"',
					"repeat with addinWorkbook in (get workbooks)",
					"try",
					"if (full name of addinWorkbook as text) is addinPath then close addinWorkbook saving no",
					"end try",
					"end repeat",
					"end tell",
					"end run"
				].join("\n"),
				"--",
				addinPath
			],
			{ stdio: "inherit" }
		);
	}
}
