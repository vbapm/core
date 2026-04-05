import { copy, readFile, writeFile } from "fs-extra";
import { join } from "path";
import { dev, empty, json, single, standard, targetless } from "./__fixtures__";
import { execute, readdir, run, RunResult, setup, tmp } from "./__helpers__/execute";

jest.setTimeout(120000);

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
});

describe("update", () => {
	test("update --target writes src changes into the specified built target", async () => {
		await setup(dev, "update-target", async cwd => {
			// 1. Build the project so a built target exists
			await execute(cwd, "build");

			// 2. Modify a [src] module
			await writeFile(
				join(cwd, "src/Validation.bas"),
				`Attribute VB_Name = "Validation"\nPublic Function GetMarker() As String\n    GetMarker = "updated"\nEnd Function\n`,
				"utf8"
			);

			// 3. Update using explicit --target
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

			// 2. Add a marker to a [dev-src] module
			await writeFile(
				join(cwd, "src/TestModule.bas"),
				`Attribute VB_Name = "TestModule"\nPublic Sub ShouldNotAppear()\nEnd Sub\n`,
				"utf8"
			);

			// 3. Update with --release (dev-src should be skipped)
			const { stdout } = await execute(cwd, "update --release");

			// 4. Export VBA only to reveal what is actually stored in the built file
			await execute(cwd, "export --vba-only --target xlsm");

			// 5. TestModule should still have the pre-update content because it was excluded
			const testModuleContent = await readFile(join(cwd, "src/TestModule.bas"), "utf8");
			expect(testModuleContent).not.toContain("ShouldNotAppear");
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
