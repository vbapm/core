/**
 * End-to-end tests for `vba update`.
 */

import { readFile, writeFile } from "fs-extra";
import { join } from "path";
import { dev, targetless } from "./__fixtures__";
import { execute, setup } from "./__helpers__/execute";

jest.setTimeout(180000);

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
