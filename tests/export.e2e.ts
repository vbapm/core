/**
 * End-to-end tests for `vba export`.
 */

import { copy, pathExists, readFile, writeFile } from "fs-extra";
import { join } from "path";
import { empty, json, standard } from "./__fixtures__";
import { execute as runCommand, readdir, setup, stripWarnings, tmp } from "./__helpers__/execute";
import { addFileMapSnapshotSerializer } from "./__helpers__/snapshot";
import { measure, timedTest } from "./__helpers__/timing";

jest.setTimeout(180000);

addFileMapSnapshotSerializer();

function execute(cwd: string, command: string) {
	return measure(`command ${command}`, () => runCommand(cwd, command));
}

describe("export", () => {
	timedTest("export to empty project", async () => {
		await setup(empty, "export-empty", async cwd => {
			await setup(standard, "export-standard", async built => {
				// 1. Build standard project
				await execute(built, "build");

				// 2. Copy standard built into empty
				await copy(join(built, "build/standard.xlsm"), join(cwd, "build/empty.xlsm"));

				// 3. Export "empty" project
				const { stdout } = await execute(cwd, "export --target xlsm");

				const result = await readdir(cwd);
				expect(result).toMatchSnapshot();
				expect(stripWarnings(stdout)).toMatchSnapshot();
			});
		});
	});

	timedTest("export to project with dependency", async () => {
		await setup(json, "export-json", async cwd => {
			await setup(standard, "export-standard-to-json", async built => {
				await execute(built, "build");

				await copy(join(built, "build/standard.xlsm"), join(cwd, "build/json.xlsm"));

				const { stdout } = await execute(cwd, "export --target xlsm");

				const result = await readdir(cwd);
				expect(result).toMatchSnapshot();
				expect(stripWarnings(stdout)).toMatchSnapshot();
			});
		});
	});

	timedTest("export fails when --xml-only and --vba-only are both specified", async () => {
		await tmp("export-options-conflict", async cwd => {
			await expect(
				execute(cwd, "export --target xlsm --xml-only --vba-only")
			).rejects.toMatchObject({
				stderr: expect.stringContaining("--xml-only and --vba-only are mutually exclusive.")
			});
		});
	});

	timedTest("export skips empty document objects when include-empty-objects = false", async () => {
		await tmp("export-no-empty-objs", async cwd => {
			// 1. Create a blank xlsm project (empty Sheet1 + ThisWorkbook)
			await execute(cwd, "new blank.xlsm");
			const dir = join(cwd, "blank");

			// 2. Add include-empty-objects = false
			let toml = await readFile(join(dir, "vbaproject.toml"), "utf-8");
			toml = toml.replace("[source]", "[source]\ninclude-empty-objects = false");
			await writeFile(join(dir, "vbaproject.toml"), toml);

			// 3. Export — empty document objects should be skipped
			await execute(dir, "export --target xlsm");

			// 4. Verify empty objects are NOT present
			const objDir = join(dir, "src", "Excel Objects");
			await expect(pathExists(join(objDir, "Sheet1.cls"))).resolves.toBe(false);
			await expect(pathExists(join(objDir, "ThisWorkbook.cls"))).resolves.toBe(false);
		});
	});
});
