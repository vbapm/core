/**
 * End-to-end tests for `vba config`.
 */

import { copy, pathExists, readFile } from "fs-extra";
import { join } from "path";
import { execute, tmp } from "./__helpers__/execute";

jest.setTimeout(180000);

describe("config", () => {
	test("writes a global background setting to the install config", async () => {
		await tmp("config-global-background", async cwd => {
			const installDir = join(cwd, "install");
			const binDir = join(installDir, "bin");

			await copy(join(process.cwd(), "bin"), binDir);
			await copy(join(process.cwd(), "lib"), join(installDir, "lib"));

			await execute(cwd, "config --global background true", { binDir, forceSubprocess: true });

			const configFile = join(binDir, "vbapm.toml");
			expect(await pathExists(configFile)).toBe(true);
			expect(await readFile(configFile, "utf8")).toMatch(/background\s*=\s*true/);
		});
	});
});
