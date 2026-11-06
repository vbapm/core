import { EventEmitter } from "events";
import { jest } from "@jest/globals";
import type { Args } from "mri";

await jest.unstable_mockModule("child_process", () => ({
	execFile: jest.fn(),
	execSync: jest.fn(),
	spawn: jest.fn()
}));

await jest.unstable_mockModule("open", () => ({
	default: jest.fn()
}));

const { spawn } = await import("child_process");
const { default: open } = await import("open");
const { env } = await import("../src/env");
const { checkTargetOpen, openTarget } = await import("../src/actions/open-target");
const { default: openCommand } = await import("../src/bin/vbapm-open");
const { resolve } = await import("../src/utils/path");
const { parse: parseToml } = await import("../src/utils/toml");

describe("openTarget", () => {
	test("checks a file argument without resolving a project target", async () => {
		const originalIsWindows = env.isWindows;
		env.isWindows = true;
		const checkProcess = Object.assign(new EventEmitter(), {
			stdout: new EventEmitter(),
			stderr: new EventEmitter()
		});
		(spawn as jest.Mock).mockImplementationOnce(() => {
			setImmediate(() => {
				checkProcess.stdout.emit(
					"data",
					JSON.stringify({ instances: [], uninspected_excel_process_ids: [] })
				);
				checkProcess.emit("close", 0);
			});
			return checkProcess;
		});
		const log = jest.spyOn(console, "log").mockImplementation(() => {});

		try {
			await openCommand({ file: "test_fake_addin.xlsm", check: true } as unknown as Args);
			expect(spawn).toHaveBeenCalledWith(
				"powershell.exe",
				expect.arrayContaining(["-Path", resolve("test_fake_addin.xlsm")]),
				expect.objectContaining({ windowsHide: true })
			);
		} finally {
			log.mockRestore();
			env.isWindows = originalIsWindows;
		}
	});

	test("opens a file argument without resolving a project target", async () => {
		const originalIsWindows = env.isWindows;
		env.isWindows = true;
		const launcher = Object.assign(new EventEmitter(), { ref: jest.fn() });
		const readinessCheck = Object.assign(new EventEmitter(), {
			stdout: new EventEmitter(),
			stderr: new EventEmitter()
		});
		(open as jest.Mock).mockImplementationOnce(async () => {
			setImmediate(() => launcher.emit("close", 0));
			return launcher;
		});
		(spawn as jest.Mock).mockImplementationOnce(() => {
			setImmediate(() => readinessCheck.emit("close", 0));
			return readinessCheck;
		});
		const log = jest.spyOn(console, "log").mockImplementation(() => {});

		try {
			await openCommand({ file: "test_fake_addin.xlsm" } as unknown as Args);
			expect(open).toHaveBeenCalledWith(resolve("test_fake_addin.xlsm"), { wait: false });
		} finally {
			(open as jest.Mock).mockClear();
			log.mockRestore();
			env.isWindows = originalIsWindows;
		}
	});

	test("returns a TOML report for exact and same-name workbook matches", async () => {
		const originalIsWindows = env.isWindows;
		env.isWindows = true;
		const report = JSON.stringify({
			uninspected_excel_process_ids: [],
			instances: [
				{
					path: "build/demo.xlsm",
					same_path: true,
					source: "workbook",
					visible: true,
					process_id: 123,
					window_handle: 456,
					access: "read_write"
				},
				{
					path: "C:/other/demo.xlsm",
					same_path: false,
					source: "workbook",
					visible: false,
					process_id: 789,
					window_handle: 101,
					access: "read_only"
				}
			]
		});
		const checkProcess = Object.assign(new EventEmitter(), {
			stdout: new EventEmitter(),
			stderr: new EventEmitter()
		});
		(spawn as jest.Mock).mockImplementationOnce(() => {
			setImmediate(() => {
				checkProcess.stdout.emit("data", report);
				checkProcess.emit("close", 0);
			});
			return checkProcess;
		});

		try {
			const output = await checkTargetOpen("build/demo.xlsm");
			const parsed = await parseToml(output);
			expect(parsed.target).toEqual({
				path: "build/demo.xlsm",
				name: "demo.xlsm",
				open: true,
				same_name_elsewhere: true,
				match_count: 2,
				all_excel_processes_inspected: true,
				uninspected_excel_process_ids: []
			});
			expect(parsed.instances).toEqual([
				{
					file_path: "build/demo.xlsm",
					match: "exact_path",
					source: "workbook",
					process_id: 123,
					window_handle: 456,
					visible: true,
					access: "read_write"
				},
				{
					file_path: "C:/other/demo.xlsm",
					match: "same_name_different_location",
					source: "workbook",
					process_id: 789,
					window_handle: 101,
					visible: false,
					access: "read_only"
				}
			]);
			expect(spawn).toHaveBeenCalledWith(
				"powershell.exe",
				expect.arrayContaining(["-File", expect.stringContaining("check-open.ps1")]),
				expect.objectContaining({ windowsHide: true })
			);
			expect(open).not.toHaveBeenCalled();
		} finally {
			env.isWindows = originalIsWindows;
		}
	});

	test("marks the report incomplete when an Excel process cannot be inspected", async () => {
		const originalIsWindows = env.isWindows;
		env.isWindows = true;
		const checkProcess = Object.assign(new EventEmitter(), {
			stdout: new EventEmitter(),
			stderr: new EventEmitter()
		});
		(spawn as jest.Mock).mockImplementationOnce(() => {
			setImmediate(() => {
				checkProcess.stdout.emit(
					"data",
					JSON.stringify({ instances: [], uninspected_excel_process_ids: [456] })
				);
				checkProcess.emit("close", 0);
			});
			return checkProcess;
		});

		try {
			const parsed = await parseToml(await checkTargetOpen("build/demo.xlsm"));
			expect(parsed.target).toEqual({
				path: "build/demo.xlsm",
				name: "demo.xlsm",
				open: false,
				same_name_elsewhere: false,
				match_count: 0,
				all_excel_processes_inspected: false,
				uninspected_excel_process_ids: [456]
			});
			expect(parsed.instances).toEqual([]);
		} finally {
			env.isWindows = originalIsWindows;
		}
	});

	test("launches the target without waiting for the application to exit", async () => {
		const launcher = Object.assign(new EventEmitter(), { ref: jest.fn() });
		const readinessCheck = Object.assign(new EventEmitter(), {
			stdout: new EventEmitter(),
			stderr: new EventEmitter()
		});
		(open as jest.Mock).mockImplementationOnce(async () => {
			setImmediate(() => launcher.emit("close", 0));
			return launcher;
		});
		(spawn as jest.Mock).mockImplementationOnce(() => {
			setImmediate(() => readinessCheck.emit("close", 0));
			return readinessCheck;
		});

		await expect(openTarget("build/demo.xlsm")).resolves.toBe(true);
		expect(open).toHaveBeenCalledWith("build/demo.xlsm", { wait: false });
		expect(launcher.ref).toHaveBeenCalled();
		if (process.platform === "win32") {
			expect(spawn).toHaveBeenCalledWith(
				"powershell.exe",
				expect.arrayContaining(["-File", expect.stringContaining("wait-open.ps1")]),
				expect.objectContaining({ windowsHide: true })
			);
		}
	});

	test("reports when Excel is unavailable to open the target", async () => {
		if (process.platform !== "win32") return;

		const launcher = Object.assign(new EventEmitter(), { ref: jest.fn() });
		const readinessCheck = Object.assign(new EventEmitter(), {
			stdout: new EventEmitter(),
			stderr: new EventEmitter()
		});
		(open as jest.Mock).mockImplementationOnce(async () => {
			setImmediate(() => launcher.emit("close", 0));
			return launcher;
		});
		(spawn as jest.Mock).mockImplementationOnce(() => {
			setImmediate(() => {
				readinessCheck.stderr.emit("data", "Excel is not available to open the target.");
				readinessCheck.emit("close", 1);
			});
			return readinessCheck;
		});

		await expect(openTarget("build/demo.xlsm")).rejects.toThrow(
			"Excel is not available to open the target."
		);
	});
});
