import { mkdtemp, mkdir, rm, writeFile } from "fs/promises";
import { tmpdir } from "os";
import { join } from "path";
import { complex, dev, empty, nameMismatch } from "../../../tests/__fixtures__";
import { reset, setup } from "../../../tests/__helpers__/project";
import { env } from "../../env";
import { loadFromProject } from "../load-from-project";
import { loadPowerQueries } from "../load-power-queries";
import { normalizeBuildGraph } from "../__helpers__/build-graph";

afterAll(reset);

test("should load BuildGraph from standard project", async () => {
	const { project, dependencies } = await setup(complex);
	const graph = await loadFromProject(project, dependencies);

	expect(normalizeBuildGraph(graph)).toMatchSnapshot();
});

test("should load BuildGraph from empty project", async () => {
	const { project, dependencies } = await setup(empty);
	const graph = await loadFromProject(project, dependencies);

	expect(normalizeBuildGraph(graph)).toMatchSnapshot();
});

test("should ignore Power Query wildcard matches in the VBA build graph", async () => {
	const { project, dependencies } = await setup(empty);
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-build-"));
	const sourceDirectory = join(root, "src", "Power Queries");

	try {
		await mkdir(sourceDirectory, { recursive: true });
		await writeFile(join(sourceDirectory, "Query.pq"), "not a VBA component", "utf8");
		project.paths.dir = root;
		project.manifest.src = [{ name: "PowerQueries", path: "src/Power Queries/*.pq" }];

		const graph = await loadFromProject(project, dependencies);
		expect(graph.components).toHaveLength(0);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});

test("loads project Power Queries separately from VBA components", async () => {
	const { project, dependencies } = await setup(empty);
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-build-"));

	try {
		await mkdir(join(root, "src", "Power Queries"), { recursive: true });
		await mkdir(join(root, "src", "manual"), { recursive: true });
		await mkdir(join(root, "dev-src"), { recursive: true });
		await writeFile(join(root, "src", "Power Queries", "Query.pq"), "let A = 1 in A", "utf8");
		await writeFile(join(root, "src", "Power Queries", "Nested.pq"), "let B = 2 in B", "utf8");
		await writeFile(join(root, "src", "manual", "Stored.pq"), "let D = 4 in D", "utf8");
		await writeFile(join(root, "dev-src", "DevQuery.pq"), "let C = 3 in C", "utf8");
		project.paths.dir = root;
		project.manifest.src = [
			{ name: "Mapped", path: join(root, "src", "manual", "Stored.pq") },
			{ name: "PowerQueries", path: "src/Power Queries/*.pq" }
		];
		project.manifest.devSrc = [{ name: "DevQuery", path: join(root, "dev-src", "DevQuery.pq") }];

		const graph = await loadFromProject(project, dependencies);
		expect(graph.components).toHaveLength(0);
		expect(graph.powerQueries).toHaveLength(4);
		expect(graph.powerQueries).toEqual(
			expect.arrayContaining([
				{ name: "Mapped", formula: "let D = 4 in D" },
				{ name: "Query", formula: "let A = 1 in A" },
				{ name: "Nested", formula: "let B = 2 in B" },
				{ name: "DevQuery", formula: "let C = 3 in C" }
			])
		);

		const releaseGraph = await loadFromProject(project, dependencies, { release: true });
		expect(releaseGraph.powerQueries).toHaveLength(3);
		expect(releaseGraph.powerQueries).toEqual(
			expect.arrayContaining([
				{ name: "Mapped", formula: "let D = 4 in D" },
				{ name: "Query", formula: "let A = 1 in A" },
				{ name: "Nested", formula: "let B = 2 in B" }
			])
		);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});

test("does not load a wildcard query already covered by an individual mapping", async () => {
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-overlap-"));
	const queryPath = join(root, "src", "pq", "Stored.pq");
	const project = {
		paths: { dir: root },
		manifest: {
			src: [
				{ name: "CustomQuery", path: queryPath },
				{ name: "PowerQueries", path: join(root, "src", "pq", "*.pq") }
			],
			devSrc: []
		}
	} as unknown as import("../../project").Project;

	try {
		await mkdir(join(root, "src", "pq"), { recursive: true });
		await writeFile(queryPath, "let Source = 1 in Source", "utf8");

		await expect(loadPowerQueries(project)).resolves.toEqual([
			{ name: "CustomQuery", formula: "let Source = 1 in Source" }
		]);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});

test("filters Power Query matches from broad source and dev-source wildcards", async () => {
	const { project, dependencies } = await setup(empty);
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-wildcard-"));

	try {
		await mkdir(join(root, "src"), { recursive: true });
		await mkdir(join(root, "dev-src"), { recursive: true });
		await writeFile(join(root, "src", "Main.bas"), 'Attribute VB_Name = "Main"\n', "utf8");
		await writeFile(join(root, "src", "Query.pq"), "not a VBA component", "utf8");
		await writeFile(join(root, "dev-src", "Dev.bas"), 'Attribute VB_Name = "Dev"\n', "utf8");
		await writeFile(join(root, "dev-src", "DevQuery.pq"), "not a VBA component", "utf8");
		project.paths.dir = root;
		project.manifest.src = [{ name: "AllSources", path: "src/**/*" }];
		project.manifest.devSrc = [{ name: "AllDevSources", path: "dev-src/**/*" }];

		const graph = await loadFromProject(project, dependencies);
		expect(graph.components.map(component => component.name).sort()).toEqual(["Dev", "Main"]);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});

test("should load BuildGraph with devDependencies", async () => {
	const { project, dependencies } = await setup(dev);
	const graph = await loadFromProject(project, dependencies);

	expect(normalizeBuildGraph(graph)).toMatchSnapshot();
});

test("should ignore dev-* for --release", async () => {
	const { project, dependencies } = await setup(dev);
	const graph = await loadFromProject(project, dependencies, { release: true });

	expect(normalizeBuildGraph(graph)).toMatchSnapshot();
});

test("should use codename for BuildGraph name when set", async () => {
	const { project, dependencies } = await setup(complex);
	project.manifest.codename = "MyProjectName";
	const graph = await loadFromProject(project, dependencies);

	expect(graph.name).toBe("MyProjectName");
});

test("should default BuildGraph name to VBAProject when codename not set", async () => {
	const { project, dependencies } = await setup(complex);
	project.manifest.codename = undefined;
	const graph = await loadFromProject(project, dependencies);

	expect(graph.name).toBe("VBAProject");
});

test("should warn when [source.files] key does not match Attribute VB_Name", async () => {
	const log = jest.spyOn(env.reporter, "log");
	const { project, dependencies } = await setup(nameMismatch);

	const graph = await loadFromProject(project, dependencies);

	// Build should still succeed despite the mismatch
	expect(graph.components).toHaveLength(1);
	expect(graph.components[0].name).toBe("Bonjour");

	// Warning should be emitted
	expect(log).toHaveBeenCalledWith(
		expect.any(String), // Message.SourceNameMismatch
		expect.stringContaining(
			'"Hello" in [source.files] does not match Attribute VB_Name = "Bonjour"'
		)
	);

	log.mockRestore();
});
