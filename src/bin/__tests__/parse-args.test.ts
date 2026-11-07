import { mkdtemp, rm } from "fs/promises";
import { tmpdir } from "os";
import { join } from "node:path";
import { loadToolSettings } from "../../config";
import { env } from "../../env";
import { parseCliArgs } from "../parse-args";
import configCommand from "../vbapm-config";

test("saves the global background setting from the reported command", async () => {
	const root = await mkdtemp(join(tmpdir(), "vbapm-global-config-"));
	const originalBin = env.bin;
	const originalCwd = env.cwd;
	const log = jest.spyOn(console, "log").mockImplementation(() => {});

	try {
		env.bin = root;
		env.cwd = root;
		const args = parseCliArgs(["config", "--global", "background", "true"]);
		args._.shift();

		await configCommand(args);

		expect(await loadToolSettings({ global: true })).toEqual({ background: true });
	} finally {
		log.mockRestore();
		env.bin = originalBin;
		env.cwd = originalCwd;
		await rm(root, { recursive: true, force: true });
	}
});

test("keeps config key and value positional after --global", () => {
	const args = parseCliArgs(["config", "--global", "background", "true"]);

	expect(args._).toEqual(["config", "background", "true"]);
	expect(args.global).toBe(true);
});

test("parses init background as a switch", () => {
	const args = parseCliArgs(["init", "--background", "--target", "xlam"]);

	expect(args.background).toBe(true);
	expect(args.target).toBe("xlam");
});
