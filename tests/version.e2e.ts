/**
 * End-to-end tests for `vba version`.
 */

import { join } from "path";
import { execute, readdir, tmp } from "./__helpers__/execute";
import { addFileMapSnapshotSerializer } from "./__helpers__/snapshot";

jest.setTimeout(180000);

addFileMapSnapshotSerializer();

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
