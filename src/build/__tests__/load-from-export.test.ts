import { mkdtemp, mkdir, rm, writeFile } from "fs/promises";
import { tmpdir } from "os";
import { join } from "path";
import { devExport, standardExport } from "../../../tests/__fixtures__";
import { reset, setupEnvironment } from "../../../tests/__helpers__/project";
import { loadFromExport } from "../load-from-export";
import { normalizeBuildGraph } from "../__helpers__/build-graph";

afterAll(reset);

test("should load BuildGraph from exported project", async () => {
	setupEnvironment(standardExport);

	const graph = await loadFromExport(standardExport);
	expect(normalizeBuildGraph(graph)).toMatchSnapshot();
});

test("should load dev-src from exported project", async () => {
	setupEnvironment(devExport);

	const graph = await loadFromExport(devExport);
	expect(normalizeBuildGraph(graph)).toMatchSnapshot();
});

test("should ignore staged Power Query files", async () => {
	setupEnvironment(standardExport);
	const staging = await mkdtemp(join(tmpdir(), "vbapm-pq-export-"));

	try {
		await mkdir(join(staging, "pq"));
		await writeFile(join(staging, "pq", "Query.pq"), "let Source = 1 in Source", "utf8");

		const graph = await loadFromExport(staging);
		expect(graph.components).toHaveLength(0);
	} finally {
		await rm(staging, { recursive: true, force: true });
	}
});
