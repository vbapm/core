import { EventEmitter } from "events";
import open from "open";
import { openTarget } from "../src/actions/open-target";

jest.mock("open", () => ({
	__esModule: true,
	default: jest.fn()
}));

describe("openTarget", () => {
	test("launches the target without waiting for the application to exit", async () => {
		const launcher = Object.assign(new EventEmitter(), { ref: jest.fn() });
		(open as jest.Mock).mockImplementationOnce(async () => {
			setImmediate(() => launcher.emit("close", 0));
			return launcher;
		});

		await expect(openTarget("build/demo.xlsm")).resolves.toBe(true);
		expect(open).toHaveBeenCalledWith("build/demo.xlsm", { wait: false });
		expect(launcher.ref).toHaveBeenCalled();
	});
});
