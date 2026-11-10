/**
 * End-to-end tests for `vba export`.
 */

import { copy, pathExists, readFile, remove, writeFile } from "fs-extra";
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
	timedTest("extract creates a missing target directory", async () => {
		await setup(standard, "extract-missing-target", async cwd => {
			await execute(cwd, "build");
			await remove(join(cwd, "targets/xlsm"));

			await execute(cwd, "extract --target xlsm --background");

			await expect(pathExists(join(cwd, "targets/xlsm/xl/workbook.xml"))).resolves.toBe(true);
		});
	});

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

	timedTest("build and extract round-trip Unicode apostrophe comments", async () => {
		await setup(standard, "export-html-entity-comments", async cwd => {
			const manifestPath = join(cwd, "vbaproject.toml");
			const manifest = await readFile(manifestPath, "utf-8");
			const asciiManifest = manifest
				.replace(
					'target = { type = "xlsm", path = "targets/xlsm" }',
					'target = { type = "xlsm", path = "targets/xlsm", encoding = "ascii" }'
				)
				.replace("\n[src]\n", '\n[source]\nencoding = "utf-8"\n\n[src]\n');

			expect(asciiManifest).toContain('encoding = "ascii"');
			expect(asciiManifest).toContain('encoding = "utf-8"');
			await writeFile(manifestPath, asciiManifest, "utf-8");

			const componentPath = join(cwd, "src/Validation.bas");
			const originalCode = await readFile(componentPath, "utf-8");
			const comment = "Café ≂̸ 한 😀";
			const separator = originalCode.endsWith("\n") ? "" : "\n";
			await writeFile(componentPath, `${originalCode}${separator}' ${comment}\n`, "utf-8");

			await execute(cwd, "build");
			await execute(cwd, "export --target xlsm");

			const extractedCode = await readFile(componentPath, "utf-8");
			expect(extractedCode).toContain(`' ${comment}`);
		});
	});

	timedTest("build and extract round-trip comments with CP1252-only substitutions", async () => {
		await setup(standard, "export-cp1252-entity-comments", async cwd => {
			const manifestPath = join(cwd, "vbaproject.toml");
			const manifest = await readFile(manifestPath, "utf-8");
			const cp1252Manifest = manifest
				.replace(
					'target = { type = "xlsm", path = "targets/xlsm" }',
					'target = { type = "xlsm", path = "targets/xlsm", encoding = "cp1252" }'
				)
				.replace("\n[src]\n", '\n[source]\nencoding = "utf-8"\n\n[src]\n');

			expect(cp1252Manifest).toContain('encoding = "cp1252"');
			expect(cp1252Manifest).toContain('encoding = "utf-8"');
			await writeFile(manifestPath, cp1252Manifest, "utf-8");

			const componentPath = join(cwd, "src/Validation.bas");
			const originalCode = await readFile(componentPath, "utf-8");
			const comment = "Café ≂̸ 한 😀";
			const separator = originalCode.endsWith("\n") ? "" : "\n";
			await writeFile(componentPath, `${originalCode}${separator}' ${comment}\n`, "utf-8");

			await execute(cwd, "build");
			await execute(cwd, "export --target xlsm");

			const extractedCode = await readFile(componentPath, "utf-8");
			expect(extractedCode).toContain(`' ${comment}`);
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
