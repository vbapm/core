/**
 * End-to-end tests for `vba new`.
 */

import { join } from "path";
import { standard } from "./__fixtures__";
import { execute, readdir, setup, tmp } from "./__helpers__/execute";
import { addFileMapSnapshotSerializer } from "./__helpers__/snapshot";

jest.setTimeout(180000);

addFileMapSnapshotSerializer();

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
