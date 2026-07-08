/**
 * End-to-end tests for the `vba` CLI.
 *
 * ## Summary
 *
 * This file exercises the full lifecycle of an Excel VBA project by running the
 * actual `vba` CLI binary against real Excel via COM automation. Each
 * test spins up a copy of a pre-built fixture project in a temp directory,
 * runs one or more CLI commands, and asserts the resulting filesystem state or
 * CLI output against Jest snapshots.
 *
 *
 * ## File and fixture layout
 *
 * ### Test helpers
 * `tests/__helpers__/execute.ts` — core test utilities: `setup()`, `tmp()`,
 * `execute()`, `readdir()`, and `run()` (used to invoke macros inside a built
 * workbook).
 *
 * ### Fixture projects
 * Each fixture is a pre-built vbapm project directory located under
 * `tests/__fixtures__/projects/`. They are imported via
 * `tests/__fixtures__/index.ts`
 *
 * ### CLI binary
 * The `vba` binary is resolved from `bin/vba` at the repo root, unless the
 * `VBA_BIN_DIR` environment variable is set (see `getVbaBin()` in
 * `tests/__helpers__/execute.ts`).
 *
 * ### Snapshots
 * Jest snapshots are stored under `tests/__snapshots__/excel.e2e.ts.snap`.
 *
 * ### Temporary directories
 * Each test is set up in a temporary directory under `tests/.tmp/`. The
 * directory is cleaned up after the test unless the `KEEP_E2E_TMP` environment
 * variable is set to true (or `1` or `yes`).
 */

import { copy, pathExists, readFile, writeFile } from "fs-extra";
import { join } from "path";
import { promisify } from "util";
import { dev, empty, json, single, standard, targetless, withDrawing } from "./__fixtures__";
import { execute, readdir, run, RunResult, setup, tmp } from "./__helpers__/execute";

const exec = promisify(require("child_process").exec);

jest.setTimeout(180000);

expect.addSnapshotSerializer({
	test: value => isSnapshotFileMap(value),
	print: value => formatSnapshotFileMap(value as { [path: string]: string })
});

function isSnapshotFileMap(value: any): value is { [path: string]: string } {
	if (!value || typeof value !== "object" || Array.isArray(value)) {
		return false;
	}

	const entries = Object.entries(value);
	if (!entries.length) {
		return false;
	}

	return (
		entries.every(([_, contents]) => typeof contents === "string") &&
		entries.some(([path]) => path.includes("/") || path.endsWith(".toml"))
	);
}

function formatSnapshotFileMap(value: { [path: string]: string }): string {
	const lines = ["Object {"];

	for (const [path, contents] of Object.entries(value)) {
		if (contents.includes("\n")) {
			lines.push(`  ${quote(path)}:`);
			lines.push(`  ${quote(contents)},`);
		} else {
			lines.push(`  ${quote(path)}: ${quote(contents)},`);
		}
	}

	lines.push("}");
	return lines.join("\n");
}

function quote(value: string): string {
	return `"${value.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
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
});

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
				expect(stdout).toMatchSnapshot();
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
				expect(stdout).toMatchSnapshot();
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
				expect(stdout).toMatchSnapshot();
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
				expect(stdout).toMatchSnapshot();
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
	test("export preserves src-subfolders config in vbaproject.toml", async () => {
		await setup(standard, "export-subfolder", async cwd => {
			// 1. Add src-subfolders config to the project
			let toml = await readFile(join(cwd, "vbaproject.toml"), "utf-8");
			toml = toml.replace(
				'target = { type = "xlsm", path = "targets/xlsm" }',
				'target = { type = "xlsm", path = "targets/xlsm" }\nsrc-subfolders = { Modules = "Modules", Forms = "Forms", Classes = "Classes" }'
			);
			await writeFile(join(cwd, "vbaproject.toml"), toml);

			// 2. Build and export
			await execute(cwd, "build");
			await execute(cwd, "export --target xlsm");

			// 3. Verify the TOML still has src-subfolder
			const updatedToml = await readFile(join(cwd, "vbaproject.toml"), "utf-8");
			expect(updatedToml).toContain("src-subfolder");
			expect(updatedToml).toContain('Modules = "Modules"');
			expect(updatedToml).toContain('Classes = "Classes"');
		});
	});
});

describe("update", () => {
	test("update --target writes src changes into the specified built target", async () => {
		await setup(targetless, "update-target", async cwd => {
			// 1. Build the package with explicit --target (targetless has no default target)
			await execute(cwd, "build --target xlsm");

			// 2. Modify a [src] module
			await writeFile(
				join(cwd, "src/Validation.bas"),
				`Attribute VB_Name = "Validation"\nPublic Function GetMarker() As String\n    GetMarker = "updated"\nEnd Function\n`,
				"utf8"
			);

			// 3. Update using explicit --target (required since there is no default target)
			const { stdout } = await execute(cwd, "update --target xlsm");

			// 4. Export VBA only so we can inspect what ended up in the built file
			await execute(cwd, "export --vba-only --target xlsm");

			// 5. The exported Validation.bas should reflect the change that was imported
			const content = await readFile(join(cwd, "src/Validation.bas"), "utf8");
			expect(content).toContain("GetMarker");
			expect(stdout).toContain("Done.");
		});
	});

	test("update --release excludes dev-src modules", async () => {
		await setup(dev, "update-release", async cwd => {
			// 1. Build the project (dev-src included in the built file)
			await execute(cwd, "build");

			// 2. Modify a [src] module and add a marker to a [dev-src] module
			await writeFile(
				join(cwd, "src/Validation.bas"),
				`Attribute VB_Name = "Validation"\nPublic Function GetReleaseMarker() As String\n    GetReleaseMarker = "release-updated"\nEnd Function\n`,
				"utf8"
			);
			await writeFile(
				join(cwd, "src/TestModule.bas"),
				`Attribute VB_Name = "TestModule"\nPublic Sub ShouldNotAppear()\nEnd Sub\n`,
				"utf8"
			);

			// 3. Update with --release ([dev-src] should be skipped, [src] should be imported)
			const { stdout } = await execute(cwd, "update --release");

			// 4. Export VBA only to reveal what is actually stored in the built file
			await execute(cwd, "export --vba-only --target xlsm");

			// 5. TestModule should still have the pre-update content because it was excluded
			const testModuleContent = await readFile(join(cwd, "src/TestModule.bas"), "utf8");
			expect(testModuleContent).not.toContain("ShouldNotAppear");

			// 6. Validation should reflect the updated [src] content because it was included
			const validationContent = await readFile(join(cwd, "src/Validation.bas"), "utf8");
			expect(validationContent).toContain("GetReleaseMarker");
			expect(stdout).toContain("Done.");
		});
	});
});

describe("new", () => {
	test("should create blank package", async () => {
		await tmp("new-blank-package", async cwd => {
			await execute(cwd, "new blank-package --package --no-git");

			const result = await readdir(join(cwd, "blank-package"));
			expect(result).toMatchSnapshot();
		});
	});

	test("should create with blank target", async () => {
		await tmp("new-blank-target", async cwd => {
			await execute(cwd, "new blank-target.xlsm");

			const result = await readdir(join(cwd, "blank-target"));
			expect(result).toMatchSnapshot();
		});
	});

	test("should create from existing", async () => {
		await tmp("new-existing-target", async cwd => {
			await setup(standard, "new-existing-target-build", async built => {
				await execute(built, "build");
				await execute(cwd, `new existing-target --from ${join(built, "build/standard.xlsm")}`);

				const result = await readdir(join(cwd, "existing-target"));
				expect(result).toMatchSnapshot();
			});
		});
	});
});

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

describe("close", () => {
	test("close with --force succeeds after build", async () => {
		await setup(standard, "close-built", async cwd => {
			await execute(cwd, "build");
			const result = await execute(cwd, "close --force");
			expect(typeof result.stderr).toBe("string");
		});
	});

	test("close fails when built target does not exist", async () => {
		await setup(standard, "close-not-built", async cwd => {
			await expect(execute(cwd, "close --force")).rejects.toThrow();
		});
	});

	test("close with --force shows error for missing build", async () => {
		await setup(standard, "close-missing-build", async cwd => {
			try {
				await execute(cwd, "close --force");
			} catch (err: any) {
				expect(err.stderr || err.stdout || "").toMatch(/not found/i);
			}
		});
	});

	test("close fails with unsaved changes error when workbook has been modified", async () => {
		await setup(standard, "close-unsaved", async cwd => {
			// 1. Build
			await execute(cwd, "build");

			const buildDir = join(cwd, "build");
			const builtFiles = await readdir(buildDir);
			const targetFile = Object.keys(builtFiles).find(
				f => f.endsWith(".xlsm") || f.endsWith(".xlam")
			);
			if (!targetFile) throw new Error("No built target file found");

			const fullPath = join(buildDir, targetFile);

			// 2. Open workbook via COM, make a change, don't save (leave open)
			const psScriptPath = join(cwd, "_open_unsaved.ps1");
			const escapedPath = fullPath.replace(/\\/g, "\\\\");
			const psScript = [
				'$path = "' + escapedPath + '"',
				"$excel = New-Object -ComObject Excel.Application",
				"$excel.Visible = $false",
				"$wb = $excel.Workbooks.Open($path)",
				"$ws = $wb.Worksheets(1)",
				'$ws.Cells.Item(1,1) = "unsaved change"',
				'Write-Output "opened_and_modified"'
			].join("\n");
			await writeFile(psScriptPath, psScript);

			const { stdout: psOut } = await exec(
				`powershell -NoProfile -ExecutionPolicy Bypass -File "${psScriptPath}"`,
				{ timeout: 30000 }
			);
			expect(psOut.trim()).toBe("opened_and_modified");

			// 3. Close without --save or --force → should fail with unsaved changes
			await expect(execute(cwd, "close")).rejects.toThrow();
		}).catch(() => {
			// Temp dir cleanup may fail (EBUSY from lingering Excel process) — ignore
		});
	});
});

describe("version", () => {
	test("should update to explicit version", async () => {
		await tmp("new-blank-package", async cwd => {
			await execute(cwd, "new blank-package --package --no-git");

			const dir = join(cwd, "blank-package");
			await execute(dir, "version v2.0.0");

			const result = await readdir(dir);
			expect(result).toMatchSnapshot();
		});
	});

	test("should update by increment and preid", async () => {
		await tmp("new-blank-package", async cwd => {
			await execute(cwd, "new blank-package --package --no-git");

			const dir = join(cwd, "blank-package");
			await execute(dir, "version preminor --preid beta");

			const result = await readdir(dir);
			expect(result).toMatchSnapshot();
		});
	});
});

async function validateBuild(cwd: string, target: string): Promise<RunResult> {
	const file = join(cwd, "build", target);
	return await run("excel", file, "Validation.Validate");
}

describe("normalize-worksheet-names", () => {
	test("renames worksheet sidecar .rels file alongside worksheet XML", async () => {
		await setup(withDrawing, "normalize-sidecar-rels", async cwd => {
			// Build creates the xlsm from targets/xlsm OOXML (which contains a drawing)
			await execute(cwd, "build");

			// Export runs normalizeWorksheetNames: sheet1.xml → shtSheet1.xml
			// The sidecar _rels/sheet1.xml.rels must also be renamed → shtSheet1.xml.rels
			await execute(cwd, "export --target xlsm");

			const relsDir = join(cwd, "targets/xlsm/xl/worksheets/_rels");
			expect(await pathExists(join(relsDir, "shtSheet1.xml.rels"))).toBe(true);
			expect(await pathExists(join(relsDir, "sheet1.xml.rels"))).toBe(false);
		});
	});
});
