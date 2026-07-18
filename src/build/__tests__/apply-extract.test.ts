import { wildcard } from "../../../tests/__fixtures__";
import { reset, setup } from "../../../tests/__helpers__/project";
import { resolveSourceFiles } from "../apply-extract";

afterEach(reset);

describe("resolveSourceFiles", () => {
	test("resolves wildcard entries to concrete file paths", async () => {
		const { project } = await setup(wildcard);

		const map = await resolveSourceFiles(project);

		// Wildcard fixture has 6 source files across 3 wildcard groups
		// (no Forms files exist on disk)
		expect(map.size).toBe(6);

		// Each entry should have a source and a loaded component
		for (const [absPath, entry] of map) {
			expect(absPath).toBe(entry.component.details.path);
			expect(entry.source).toBeDefined();
			expect(entry.source.name).toBeDefined();
			expect(entry.source.path).toBeDefined();
			expect(entry.component.code).toBeDefined();
			expect(entry.component.name).toBeDefined();
		}

		// Verify specific files are present by name
		const names = [...map.values()].map(e => e.component.name).sort();
		expect(names).toEqual([
			"Class1",
			"Sheet1",
			"Sheet2",
			"Sheet3",
			"ThisWorkbook",
			"Validation"
		]);
	});

	test("returns empty map when manifest has no src entries", async () => {
		const { project } = await setup(wildcard);
		project.manifest.src = [];

		const map = await resolveSourceFiles(project);

		expect(map.size).toBe(0);
	});

	test("resolves single-path (non-wildcard) entries", async () => {
		const { project } = await setup(wildcard);
		// Replace wildcards with a single explicit entry
		project.manifest.src = [
			{
				name: "Validation",
				path: project.paths.dir + "/src/Modules/Validation.bas"
			}
		];

		const map = await resolveSourceFiles(project);

		expect(map.size).toBe(1);
		const entry = [...map.values()][0];
		expect(entry.source.name).toBe("Validation");
		expect(entry.component.name).toBe("Validation");
	});
});
