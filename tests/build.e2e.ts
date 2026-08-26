/**
 * End-to-end tests for `vba build`.
 *
 * Each test builds a fixture project and then validates the built workbook by
 * running a macro inside it (see `validateBuild`), snapshotting the result.
 */

import { readFile, writeFile } from "fs-extra";
import { join } from "path";
import { single, standard, targetless } from "./__fixtures__";
import { execute, closePersistentSession, run, RunResult, setup } from "./__helpers__/execute";

jest.setTimeout(180000);

// When the persistent session is enabled (VBA_PERSISTENT_SESSION=1), the
// in-process `run()` helper keeps one hidden Excel instance alive for the whole
// suite. Close it here so we don't leak it after the tests finish.
afterAll(async () => {
	await closePersistentSession();
});

async function validateBuild(cwd: string, target: string): Promise<RunResult> {
	const file = join(cwd, "build", target);
	return await run("excel", file, "Validation.Validate");
}

describe("build", () => {
	test("build standard project", async () => {
		await setup(standard, "build", async cwd => {
			await execute(cwd, "build");

			const result = await validateBuild(cwd, "standard.xlsm");
			expect(result).toMatchSnapshot();
		});
	});

	test("build project with single target", async () => {
		await setup(single, "build-single", async cwd => {
			await execute(cwd, "build");

			const result = await validateBuild(cwd, "single.xlsm");
			expect(result).toMatchSnapshot();
		});
	});

	test("build project with no target", async () => {
		await setup(targetless, "build-targetless", async cwd => {
			await execute(cwd, "build --target xlsm");

			const result = await validateBuild(cwd, "targetless.xlsm");
			expect(result).toMatchSnapshot();
		});
	});

	test("build --vba-only updates an existing target", async () => {
		await setup(targetless, "build-vba-only", async cwd => {
			await execute(cwd, "build --target xlsm");

			await writeFile(
				join(cwd, "src/Validation.bas"),
				`Attribute VB_Name = "Validation"\nPublic Function GetMarker() As String\n    GetMarker = "build-vba-only"\nEnd Function\n`,
				"utf8"
			);

			const { stdout } = await execute(cwd, "build --vba-only --target xlsm");
			expect(stdout).toContain("Updating VBA");

			await execute(cwd, "export --vba-only --target xlsm");
			const content = await readFile(join(cwd, "src/Validation.bas"), "utf8");
			expect(content).toContain('GetMarker = "build-vba-only"');
		});
	});
});
