import { mkdtemp, rm } from "fs/promises";
import { tmpdir } from "os";
import { join } from "path";
import { codepageToLabel, getSystemCodepage } from "../../build/encoding-sniffer";
import { initProject } from "../init-project";
import { writeFile } from "../../utils/fs";

test("does not add PowerQueries to new project manifests by default", async () => {
	const dir = await mkdtemp(join(tmpdir(), "vbapm-init-power-query-"));
	const writeFileSpy = jest.spyOn({ writeFile }, "writeFile");

	try {
		await initProject({
			name: "power-query-test",
			dir,
			pkg: true,
			git: false,
			configTemplates: false
		});

		const manifestWrite = writeFileSpy.mock.calls.find(([path]) =>
			String(path).endsWith("vbaproject.toml")
		);
		expect(manifestWrite?.[1]).not.toContain("PowerQueries");
		expect(manifestWrite?.[1]).toContain(`encoding = "${codepageToLabel(getSystemCodepage())}"`);
	} finally {
		writeFileSpy.mockRestore();
		await rm(dir, { recursive: true, force: true });
	}
});
