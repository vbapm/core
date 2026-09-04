/**
 * End-to-end tests for wildcard source extraction (`vba extract`).
 */

import { conflict, wildcard } from "./__fixtures__";
import { execute, readdir, setup } from "./__helpers__/execute";
import { addFileMapSnapshotSerializer } from "./__helpers__/snapshot";

jest.setTimeout(180000);

addFileMapSnapshotSerializer();

describe("wildcard extract", () => {
	test("extract does not add individual entries for wildcard-covered modules", async () => {
		await setup(wildcard, "extract-wildcard", async cwd => {
			// 1. Build wildcard project to produce build/wildcard.xlsm
			await execute(cwd, "build");

			// 2. Extract from wildcard
			await execute(cwd, "extract --target xlsm");

			// 3. Verify vbaproject.toml is unchanged (no individual entries added,
			//    wildcard patterns preserved, references intact)
			const result = await readdir(cwd);
			expect(result).toMatchSnapshot();
		});
	});

	test("extract renames mismatched manifest entries and preserves both", async () => {
		await setup(conflict, "extract-conflict", async cwd => {
			// 1. Extract from the conflict fixture
			const { stdout } = await execute(cwd, "extract --target xlsm");

			// 2. Warning should mention the rename
			expect(stdout).toContain('"Module1" in [source.files] was renamed to "Validation"');

			// 3. Verify both entries exist in vbaproject.toml:
			//    Module1 = "src/Module1.bas" (new workbook component)
			//    Validation = "src/Validation.bas" (renamed from Module1)
			const result = await readdir(cwd);
			expect(result).toMatchSnapshot();
		});
	});
});
