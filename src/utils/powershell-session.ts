import { ChildProcessWithoutNullStreams, spawn } from "child_process";
import { readFile, unlink } from "fs/promises";
import { Socket, connect } from "net";
import { join as pathJoin } from "path";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { pathExists } from "./fs";

const debug = env.debug("vbapm:ps-session");

const RECORD_SEP = "\u001e";
const UNIT_SEP = "\u001f";

interface SessionRequest {
	id: string;
	appName: string;
	file: string;
	macro: string;
	keepOpen: boolean;
	background: boolean;
	args: string[];
}

interface SessionResponse {
	success: boolean;
	result?: unknown;
	errors?: string[];
}

/** Result shape compatible with RunResult in ./run (kept local to avoid a cycle). */
export interface SessionRunResult {
	success: boolean;
	messages: string[];
	warnings: string[];
	errors: string[];
	stdout?: string;
	stderr?: string;
}

/**
 * A long-lived PowerShell process that drives a *reused* Excel.Application
 * instance across multiple macro invocations.
 *
 * Instead of spawning a fresh `powershell.exe -File run.ps1` per call (which
 * destroys the COM stub when the process exits), we spawn one PowerShell that
 * reads framed JSON requests from stdin and keeps the $app alive between them.
 */
export class PowerShellSession {
	private child: ChildProcessWithoutNullStreams | null = null;
	private socket: Socket | null = null;
	private sessionScript: string;
	private pending = new Map<
		string,
		{ resolve: (r: SessionResponse) => void; reject: (e: Error) => void }
	>();
	private buffer = "";
	private stderrBuffer = "";
	private requestSeq = 0;

	constructor(sessionScript: string) {
		this.sessionScript = sessionScript;
	}

	async start(): Promise<void> {
		if (this.child || this.socket) return;

		if (!(await pathExists(this.sessionScript))) {
			throw new CliError(
				ErrorCode.RunScriptNotFound,
				`Bridge session script not found at "${this.sessionScript}".`
			);
		}

		const isPrimaryWorker = !process.env.JEST_WORKER_ID || process.env.JEST_WORKER_ID === "1";
		const warmMetadataPath = pathJoin(
			process.env.TEMP || process.env.TMP || ".",
			"Excel-Instances",
			"warm-session.json"
		);
		if (isPrimaryWorker) {
			try {
				const metadata = JSON.parse(await readFile(warmMetadataPath, "utf8")) as { port?: number };
				const warmPort = metadata.port;
				if (typeof warmPort === "number") {
					const warmSocket = await new Promise<Socket>((resolve, reject) => {
						const candidate = connect({ host: "127.0.0.1", port: warmPort }, () =>
							resolve(candidate)
						);
						candidate.once("error", reject);
					});
					this.attachSocket(warmSocket);
					return;
				}
			} catch {
				try {
					await unlink(warmMetadataPath);
				} catch {}
			}
		}

		this.child = spawn(
			"powershell.exe",
			["-NoProfile", "-ExecutionPolicy", "Bypass", "-NoLogo", "-File", this.sessionScript],
			{
				env: process.env,
				windowsHide: true,
				stdio: ["pipe", "pipe", "pipe"]
			}
		);

		this.child.stdout.on("data", (data: Buffer) => this.onStdout(data.toString()));
		this.child.stderr.on("data", (data: Buffer) => {
			this.stderrBuffer += data.toString();
			debug("stderr:", data.toString());
		});
		this.child.on("error", err => {
			this.rejectAll(err);
			this.child = null;
		});
		this.child.on("close", code => {
			debug("session exited:", code);
			const detail = this.stderrBuffer.trim();
			this.rejectAll(
				new Error(`PowerShell session exited (code ${code})${detail ? `: ${detail}` : ""}`)
			);
			this.stderrBuffer = "";
			this.child = null;
		});
	}

	private attachSocket(socket: Socket): void {
		this.socket = socket;
		socket.on("data", (data: Buffer) => this.onStdout(data.toString()));
		socket.on("error", err => {
			this.rejectAll(err);
			this.socket = null;
		});
		socket.on("close", () => {
			this.rejectAll(new Error("Warm PowerShell session socket closed"));
			this.socket = null;
		});
	}

	async run(
		appName: string,
		file: string,
		macro: string,
		args: string[],
		options: { keepOpen?: boolean; background?: boolean } = {}
	): Promise<SessionRunResult> {
		await this.start();

		const id = `req-${++this.requestSeq}-${Date.now()}-${Math.random().toString(36).slice(2)}`;

		const request: SessionRequest = {
			id,
			appName,
			file,
			macro,
			keepOpen: !!options.keepOpen,
			background: !!options.background,
			args
		};

		// Base64-encode the JSON so args with quotes/newlines need no escaping.
		const encoded = Buffer.from(JSON.stringify(request), "utf8").toString("base64");

		return new Promise<SessionRunResult>((resolve, reject) => {
			this.pending.set(id, {
				resolve: (r: SessionResponse) => {
					resolve(this.toRunResult(r));
				},
				reject
			});

			this.writeLine(encoded);
		});
	}

	private writeLine(encoded: string): void {
		if (this.socket) this.socket.write(encoded + "\n");
		else this.child!.stdin.write(encoded + "\n");
	}

	async reset(): Promise<void> {
		if (!this.child) return;

		const id = "__VBA_RESET__";
		const encoded = Buffer.from(JSON.stringify({ id: "__VBA_RESET__" }), "utf8").toString("base64");

		await new Promise<void>((resolve, reject) => {
			this.pending.set(id, {
				resolve: (response: SessionResponse) => {
					if (response.success) {
						resolve();
					} else {
						reject(
							new Error((response.errors ?? []).join("\n") || "PowerShell session reset failed")
						);
					}
				},
				reject
			});

			this.writeLine(encoded);
		});
	}

	async close(): Promise<void> {
		if (this.socket) {
			const socket = this.socket;
			const encoded = Buffer.from(JSON.stringify({ id: "__VBA_QUIT__" }), "utf8").toString(
				"base64"
			);
			await new Promise<void>((resolve, reject) => {
				this.pending.set("__VBA_QUIT__", {
					resolve: () => {
						socket.end();
						resolve();
					},
					reject
				});
				this.writeLine(encoded);
			});
			this.socket = null;
			return;
		}
		if (!this.child) return;

		const child = this.child;

		// Drop our listeners first. Under Jest the session is closed in `afterAll`,
		// and a late `close`/`error` event would call `debug()` -> lazy `require()`
		// after the test environment is torn down, which crashes the worker.
		child.removeAllListeners("close");
		child.removeAllListeners("error");
		child.stdout.removeAllListeners("data");
		child.stderr.removeAllListeners("data");

		let timedOut = false;
		await new Promise<void>(resolve => {
			const timeout = setTimeout(() => {
				timedOut = true;
				resolve();
			}, 10000);
			child.once("close", () => {
				clearTimeout(timeout);
				resolve();
			});

			const encoded = Buffer.from(JSON.stringify({ id: "__VBA_QUIT__" }), "utf8").toString(
				"base64"
			);
			child.stdin.write(encoded + "\n");
			child.stdin.end();
		});
		if (timedOut) {
			child.kill();
		}
		this.child = null;
		session = null;
	}

	async detach(): Promise<void> {
		if (this.socket) {
			const socket = this.socket;
			const encoded = Buffer.from(JSON.stringify({ id: "__VBA_DETACH__" }), "utf8").toString(
				"base64"
			);
			await new Promise<void>((resolve, reject) => {
				this.pending.set("__VBA_DETACH__", {
					resolve: () => {
						socket.end();
						resolve();
					},
					reject
				});
				this.writeLine(encoded);
			});
			this.socket = null;
			return;
		}
		if (!this.child) return;

		const child = this.child;
		const encoded = Buffer.from(JSON.stringify({ id: "__VBA_DETACH__" }), "utf8").toString(
			"base64"
		);
		await new Promise<void>((resolve, reject) => {
			this.pending.set("__VBA_DETACH__", {
				resolve: (response: SessionResponse) => {
					if (response.success) resolve();
					else
						reject(
							new Error((response.errors ?? []).join("\n") || "PowerShell session detach failed")
						);
				},
				reject
			});
			this.writeLine(encoded);
		});

		child.removeAllListeners("close");
		child.removeAllListeners("error");
		child.stdout.removeAllListeners("data");
		child.stderr.removeAllListeners("data");
		child.unref();
		(child.stdin as any).unref?.();
		(child.stdout as any).unref?.();
		(child.stderr as any).unref?.();
		this.child = null;
		session = null;
	}

	private onStdout(chunk: string): void {
		this.buffer += chunk;

		// Responses are newline-terminated records: \u001e<id>\u001f<json>
		let nl = this.buffer.indexOf("\n");
		while (nl !== -1) {
			const line = this.buffer.slice(0, nl);
			this.buffer = this.buffer.slice(nl + 1);
			this.handleLine(line);
			nl = this.buffer.indexOf("\n");
		}
	}

	private handleLine(line: string): void {
		if (!line.startsWith(RECORD_SEP)) return;

		const sepIndex = line.indexOf(UNIT_SEP, 1);
		if (sepIndex === -1) return;

		const id = line.slice(1, sepIndex);
		const json = line.slice(sepIndex + 1);

		const pending = this.pending.get(id);
		if (!pending) return;
		this.pending.delete(id);

		try {
			pending.resolve(JSON.parse(json) as SessionResponse);
		} catch (err) {
			pending.reject(err as Error);
		}
	}

	private rejectAll(err: Error): void {
		for (const [, pending] of this.pending) {
			pending.reject(err);
		}
		this.pending.clear();
	}

	private toRunResult(r: SessionResponse): SessionRunResult {
		if (r.success) {
			if (typeof r.result === "string") {
				try {
					const parsed = JSON.parse(r.result);
					if (parsed && typeof parsed === "object" && "success" in parsed) {
						return {
							success: !!parsed.success,
							messages: parsed.messages ?? [],
							warnings: parsed.warnings ?? [],
							errors: parsed.errors ?? [],
							stdout: `${r.result}\n`,
							stderr: ""
						};
					}
				} catch {
					// Treat non-JSON macro output as ordinary stdout below.
				}
			}

			return {
				success: true,
				messages: [typeof r.result === "string" ? r.result : JSON.stringify(r.result)],
				warnings: [],
				errors: [],
				stdout: `${typeof r.result === "string" ? r.result : JSON.stringify(r.result)}\n`,
				stderr: ""
			};
		}
		return {
			success: false,
			messages: [],
			warnings: [],
			errors: r.errors ?? ["Unknown error"]
		};
	}
}

let session: PowerShellSession | null = null;

/** Singleton accessor for the process-wide persistent PowerShell session. */
export function getPowerShellSession(): PowerShellSession | null {
	if (process.platform !== "win32") return null;
	return session;
}

export function initPowerShellSession(sessionScript: string): PowerShellSession {
	if (!session) {
		session = new PowerShellSession(sessionScript);
	}
	return session;
}

export async function resetPowerShellSession(): Promise<void> {
	if (session) {
		await session.reset();
	}
}
