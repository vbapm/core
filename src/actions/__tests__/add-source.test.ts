import { mkdtemp, rm, writeFile as writeDiskFile } from "fs/promises";
import { tmpdir } from "os";
import { join } from "../../utils/path";
import { loadManifest } from "../../manifest";
import { reset, readFile } from "../../utils/fs";
import { addSource } from "../add-source";

afterEach(reset);

test("creates and registers a Power Query source file", async () => {
	const dir = await mkdtemp(join(tmpdir(), "vbapm-add-power-query-"));
	await writeDiskFile(
		join(dir, "vbaproject.toml"),
		'[project]\nname = "add-power-query-test"\ntarget = "xlsm"\n',
		"utf8"
	);

	try {
		const result = await addSource({ name: "New Query.pq", dir, dev: false });
		const expectedPath = join(dir, "src", "New Query.pq");

		expect(result).toEqual({ path: expectedPath, isNew: true });
		expect(await readFile(expectedPath, "utf8")).toBe("let\n    Source = null\nin\n    Source\n");
		expect((await loadManifest(dir)).src).toEqual([{ name: "New Query", path: expectedPath }]);
	} finally {
		await rm(dir, { recursive: true, force: true });
	}
});
