/**
 * End-to-end tests for export mode options.
 */

import { copy, readFile, writeFile } from "fs-extra";
import { join } from "path";
import { empty, standard } from "./__fixtures__";
import { execute as runCommand, readdir, setup, stripWarnings } from "./__helpers__/execute";
import { addFileMapSnapshotSerializer } from "./__helpers__/snapshot";
import { measure, timedTest } from "./__helpers__/timing";

jest.setTimeout(180000);

addFileMapSnapshotSerializer();

function execute(cwd: string, command: string) {
	return measure(`command ${command}`, () => runCommand(cwd, command));
}

describe("export modes", () => {
	timedTest("export with --xml-only skips VBA source export", async () => {
		await setup(empty, "export-xml-only", async cwd => {
			await setup(standard, "export-standard-xml-only", async built => {
				await execute(built, "build");
				await copy(join(built, "build/standard.xlsm"), join(cwd, "build/empty.xlsm"));

				const { stdout } = await execute(cwd, "export --target xlsm --xml-only");

				const result = await readdir(cwd);
				expect(result).toMatchSnapshot();
				expect(stripWarnings(stdout)).toMatchSnapshot();
			});
		});
	});

	timedTest("export with --vba-only skips XML extraction", async () => {
		await setup(empty, "export-vba-only", async cwd => {
			await setup(standard, "export-standard-vba-only", async built => {
				await execute(built, "build");
				await copy(join(built, "build/standard.xlsm"), join(cwd, "build/empty.xlsm"));

				const { stdout } = await execute(cwd, "export --target xlsm --vba-only");

				const result = await readdir(cwd);
				expect(result).toMatchSnapshot();
				expect(stripWarnings(stdout)).toMatchSnapshot();
			});
		});
	});

	timedTest("export preserves subfolders config in vbaproject.toml", async () => {
		await setup(standard, "export-subfolder", async cwd => {
			let toml = await readFile(join(cwd, "vbaproject.toml"), "utf-8");
			toml = toml.replace(
				'target = { type = "xlsm", path = "targets/xlsm" }',
				'target = { type = "xlsm", path = "targets/xlsm" }\n\n[source]\nsubfolders = { Modules = "Modules", Forms = "Forms", Classes = "Classes" }'
			);
			await writeFile(join(cwd, "vbaproject.toml"), toml);

			await execute(cwd, "build");
			await execute(cwd, "export --target xlsm");

			const updatedToml = await readFile(join(cwd, "vbaproject.toml"), "utf-8");
			expect(updatedToml).toContain("subfolders");
			expect(updatedToml).toContain('Modules = "Modules"');
			expect(updatedToml).toContain('Classes = "Classes"');
		});
	});
});
