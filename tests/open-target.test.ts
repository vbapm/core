import { EventEmitter } from "events";
import { jest } from "@jest/globals";

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
const { openTarget } = await import("../src/actions/open-target");

describe("openTarget", () => {
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
