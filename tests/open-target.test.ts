import { EventEmitter } from "events";
import { spawn } from "child_process";
import open from "open";
import { openTarget } from "../src/actions/open-target";

jest.mock("child_process", () => ({
	...jest.requireActual("child_process"),
	spawn: jest.fn()
}));

jest.mock("open", () => ({
	__esModule: true,
	default: jest.fn()
}));

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
});
