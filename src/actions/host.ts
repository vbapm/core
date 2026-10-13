import { spawn } from "child_process";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { pathExists } from "../utils/fs";
import { join } from "../utils/path";

export type HostCloseScope = "all" | "visible" | "invisible";

export async function runHostCommand(closeScope?: HostCloseScope, force = false): Promise<string> {
	if (!env.isWindows) {
		throw new CliError(
			ErrorCode.HostUnsupportedPlatform,
			"The host command is only supported on Windows."
		);
	}

	const script = join(env.scripts, "host.ps1");
	if (!(await pathExists(script))) {
		throw new CliError(ErrorCode.RunScriptNotFound, `Host script not found at "${script}".`);
	}

	const action = closeScope ? `close-${closeScope}` : "list";
	return await runPowerShell(script, ["-Action", action, ...(force ? ["-Force"] : [])]);
}

function runPowerShell(script: string, args: string[]): Promise<string> {
	return new Promise((resolve, reject) => {
		const child = spawn(
			"powershell.exe",
			["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script, ...args],
			{ windowsHide: true, stdio: ["ignore", "pipe", "pipe"] }
		);

		let stdout = "";
		let stderr = "";

		child.stdout?.on("data", data => {
			stdout += data.toString();
		});
		child.stderr?.on("data", data => {
			stderr += data.toString();
		});

		child.on("error", error =>
			reject(new CliError(ErrorCode.HostCommandFailed, error.message, error))
		);
		child.on("close", code => {
			if (code === 0) {
				resolve(stdout.trim());
				return;
			}

			const message = stderr.trim() || stdout.trim() || `PowerShell exited with code ${code}.`;
			reject(new CliError(ErrorCode.HostCommandFailed, message));
		});
	});
}
