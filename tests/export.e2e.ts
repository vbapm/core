/**
 * End-to-end tests for `vba export`.
 */

import { copy, pathExists, readFile, writeFile } from "fs-extra";
import { join } from "path";
import { empty, json, standard } from "./__fixtures__";
import { execute, readdir, setup, stripWarnings, tmp } from "./__helpers__/execute";
import { addFileMapSnapshotSerializer } from "./__helpers__/snapshot";

jest.setTimeout(180000);

addFileMapSnapshotSerializer();

describe("export", () => {
	test("export to empty project", async () => {
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

	test("export to project with dependency", async () => {
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

	test("export with --xml-only skips VBA source export", async () => {
		await setup(empty, "export-xml-only", async cwd => {
			await setup(standard, "export-standard-xml-only", async built => {
				// 1. Build standard project (provides a valid xlsm to extract XML from)
				await execute(built, "build");

				// 2. Copy built xlsm into empty project
				await copy(join(built, "build/standard.xlsm"), join(cwd, "build/empty.xlsm"));

				// 3. Export --xml-only: should update targets/xlsm/ but NOT create src/
				const { stdout } = await execute(cwd, "export --target xlsm --xml-only");

				const result = await readdir(cwd);
				expect(result).toMatchSnapshot();
				expect(stripWarnings(stdout)).toMatchSnapshot();
			});
		});
	});

	test("export with --vba-only skips XML extraction", async () => {
		await setup(empty, "export-vba-only", async cwd => {
			await setup(standard, "export-standard-vba-only", async built => {
				// 1. Build standard project
				await execute(built, "build");

				// 2. Copy built xlsm into empty project
				await copy(join(built, "build/standard.xlsm"), join(cwd, "build/empty.xlsm"));

				// 3. Export --vba-only: should create src/ but NOT update targets/xlsm/
				const { stdout } = await execute(cwd, "export --target xlsm --vba-only");

				const result = await readdir(cwd);
				expect(result).toMatchSnapshot();
				expect(stripWarnings(stdout)).toMatchSnapshot();
			});
		});
	});

	test("export fails when --xml-only and --vba-only are both specified", async () => {
		await tmp("export-options-conflict", async cwd => {
			await expect(
				execute(cwd, "export --target xlsm --xml-only --vba-only")
			).rejects.toMatchObject({
				stderr: expect.stringContaining("--xml-only and --vba-only are mutually exclusive.")
			});
		});
	});
	test("export preserves subfolders config in vbaproject.toml", async () => {
		await setup(standard, "export-subfolder", async cwd => {
			// 1. Add subfolders config via [source]
			let toml = await readFile(join(cwd, "vbaproject.toml"), "utf-8");
			toml = toml.replace(
				'target = { type = "xlsm", path = "targets/xlsm" }',
				'target = { type = "xlsm", path = "targets/xlsm" }\n\n[source]\nsubfolders = { Modules = "Modules", Forms = "Forms", Classes = "Classes" }'
			);
			await writeFile(join(cwd, "vbaproject.toml"), toml);

			// 2. Build and export
			await execute(cwd, "build");
			await execute(cwd, "export --target xlsm");

			// 3. Verify the TOML still has subfolders
			const updatedToml = await readFile(join(cwd, "vbaproject.toml"), "utf-8");
			expect(updatedToml).toContain("subfolders");
			expect(updatedToml).toContain('Modules = "Modules"');
			expect(updatedToml).toContain('Classes = "Classes"');
		});
	});

	test("export skips empty document objects when include-empty-objects = false", async () => {
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
