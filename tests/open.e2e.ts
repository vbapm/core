/**
 * End-to-end tests for `vba open`.
 */

import { standard } from "./__fixtures__";
import { execute, setup } from "./__helpers__/execute";

jest.setTimeout(180000);

describe("open", () => {
	test("open fails when built target does not exist", async () => {
		await setup(standard, "open-not-built", async cwd => {
			// Build not run, so no built file exists
			await expect(execute(cwd, "open")).rejects.toThrow();
		});
	});

	test("open shows error for missing build", async () => {
		await setup(standard, "open-missing-build", async cwd => {
			try {
				await execute(cwd, "open");
			} catch (err: any) {
				expect(err.stderr || err.stdout || "").toMatch(/built target/i);
			}
		});
	});
});
