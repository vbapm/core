import { mkdtemp, mkdir, rm, writeFile } from "fs/promises";
import { tmpdir } from "os";
import { env } from "../env";
import { loadEffectiveToolSettings, resolveBackgroundMode, saveToolSettings } from "../config";
import { join } from "../utils/path";

describe("tool config", () => {
	const originalBin = env.bin;
	const originalCwd = env.cwd;

	afterEach(async () => {
		env.bin = originalBin;
		env.cwd = originalCwd;
	});

	test("prefers local background settings over the global config", async () => {
		const root = await mkdtemp(join(tmpdir(), "vbapm-config-"));
		const projectDir = join(root, "project");
		const globalDir = join(root, "global");
		const manifestFile = join(projectDir, "vbaproject.toml");
		const projectFile = join(projectDir, "vba.toml");
		const globalFile = join(globalDir, "vba.toml");

		try {
			await mkdir(projectDir, { recursive: true });
			await mkdir(globalDir, { recursive: true });
			await writeFile(manifestFile, '[project]\nname = "demo"\ntarget = "xlsm"\n');

			env.bin = globalDir;
			env.cwd = projectDir;

			await saveToolSettings({ background: false }, { global: true, file: globalFile });
			await saveToolSettings({ background: true }, { file: projectFile });

			expect(await resolveBackgroundMode()).toBe(true);
			expect(await loadEffectiveToolSettings()).toEqual({ background: true });
		} finally {
			await rm(root, { recursive: true, force: true });
		}
	});
});
