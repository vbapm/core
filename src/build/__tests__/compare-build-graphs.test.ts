import {
	complex,
	dev,
	devExport,
	dir,
	standard,
	standardChangesExport,
	standardExport,
	wildcard
} from "../../../tests/__fixtures__";
import { reset, setup } from "../../../tests/__helpers__/project";
import { Changeset } from "../changeset";
import { compareBuildGraphs } from "../compare-build-graphs";
import { loadFromExport } from "../load-from-export";
import { loadFromProject } from "../load-from-project";
import { normalizeComponent } from "../__helpers__/component";

afterAll(reset);

test("should find no changes between build graphs", async () => {
	const { project, dependencies } = await setup(standard);

	const before = await loadFromProject(project, dependencies);
	const after = await loadFromExport(standardExport);

	const changeset = compareBuildGraphs(before, after);

	expect(changeset.components.added.length).toEqual(0);
	expect(changeset.components.changed.length).toEqual(0);
	expect(changeset.components.removed.length).toEqual(0);
}, 10000);

test("should find added, changed, and removed between build graphs", async () => {
	const { project, dependencies } = await setup(complex);

	const before = await loadFromProject(project, dependencies);
	const after = await loadFromExport(standardChangesExport);

	const changeset = compareBuildGraphs(before, after);
	expect(normalizeChangeset(changeset)).toMatchSnapshot();

	// TODO finds UserForm1 as changed, look into why
}, 10000);

test("should find no changes for dev-src", async () => {
	const { project, dependencies } = await setup(dev);

	const before = await loadFromProject(project, dependencies);
	const after = await loadFromExport(devExport);

	const changeset = compareBuildGraphs(before, after);
	expect(normalizeChangeset(changeset)).toMatchSnapshot();
}, 10000);

test("should find no changes with wildcard sources", async () => {
	const { project, dependencies } = await setup(wildcard);

	const before = await loadFromProject(project, dependencies);
	const after = await loadFromExport(standardExport);

	const changeset = compareBuildGraphs(before, after);

	expect(changeset.components.added.length).toEqual(0);
	expect(changeset.components.changed.length).toEqual(0);
	expect(changeset.components.removed.length).toEqual(0);
}, 10000);

test("should find added, changed, and removed with wildcard sources", async () => {
	const { project, dependencies } = await setup(wildcard);

	const before = await loadFromProject(project, dependencies);
	const after = await loadFromExport(standardChangesExport);

	const changeset = compareBuildGraphs(before, after);
	expect(normalizeChangeset(changeset)).toMatchSnapshot();
}, 10000);

export function normalizeChangeset(changeset: Changeset): Changeset {
	const { components, references } = changeset;

	const added = components.added.map(component => normalizeComponent(component, dir));
	const changed = components.changed.map(component => normalizeComponent(component, dir));
	const removed = components.removed.map(component => normalizeComponent(component, dir));

	return {
		components: { added, changed, removed },
		references
	};
}
