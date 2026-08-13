import { complex, dev, empty, nameMismatch } from "../../../tests/__fixtures__";
import { reset, setup } from "../../../tests/__helpers__/project";
import { env } from "../../env";
import { loadFromProject } from "../load-from-project";
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
