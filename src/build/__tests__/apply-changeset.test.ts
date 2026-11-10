import { complex, standardChangesExport, wildcard } from "../../../tests/__fixtures__";
import { reset, setup } from "../../../tests/__helpers__/project";
import { normalizeManifest } from "../../__helpers__/manifest";
import { applyChangeset, writeComponent } from "../apply-changeset";
import { compareBuildGraphs } from "../compare-build-graphs";
import { Component } from "../component";
import { Codepage } from "../encoding-sniffer";
import { loadFromExport } from "../load-from-export";
import { loadFromProject } from "../load-from-project";
import { writeFile, remove } from "../../utils/fs";

beforeEach(() => {
	jest.clearAllMocks();
});

afterEach(reset);

test("should apply changeset for project", async () => {
	jest.spyOn(console, "log").mockImplementation(() => {});

	const { project, dependencies } = await setup(complex, { silent: false });
	const before = await loadFromProject(project, dependencies);
	const after = await loadFromExport(standardChangesExport);
	const changeset = compareBuildGraphs(before, after);

	await applyChangeset(project, changeset);

	expect(normalizeManifest(project.manifest)).toMatchSnapshot();
}, 15000);

test("should write new files to subfolders without adding individual entries when covered by wildcards", async () => {
	jest.spyOn(console, "log").mockImplementation(() => {});

	const { project, dependencies } = await setup(wildcard, { silent: false });
	const before = await loadFromProject(project, dependencies);
	const after = await loadFromExport(standardChangesExport);
	const changeset = compareBuildGraphs(before, after);

	await applyChangeset(project, changeset);

	// New source files should be written to type-based subfolders
	const writeCalls = (writeFile as jest.Mock).mock.calls.map((call: any[]) => call[0]);
	const addedPaths = writeCalls.filter(
		(p: string) =>
			p.includes("Added.bas") ||
			p.includes("WebHelpers.bas") ||
			p.includes("UserForm1.frm") ||
			p.includes("IWebAuthenticator.cls")
	);

	expect(addedPaths).toEqual(
		expect.arrayContaining([
			expect.stringContaining("src/Modules/Added.bas"),
			expect.stringContaining("src/Modules/WebHelpers.bas"),
			expect.stringContaining("src/Forms/UserForm1.frm"),
			expect.stringContaining("src/Class Modules/IWebAuthenticator.cls")
		])
	);

	// Removed components should be deleted from disk
	expect(remove).toHaveBeenCalledWith(expect.stringContaining("src/Modules/Validation.bas"));

	// Manifest should keep only the 4 wildcard entries — no individual
	// entries for components already covered by wildcard patterns
	expect(normalizeManifest(project.manifest)).toMatchSnapshot();
});

test("writes a UTF-8 BOM for source encoding utf-8-bom", async () => {
	const path = "src/Module.bas";
	const code = 'Attribute VB_Name = "Module"\r\n\' café\r\n';
	const component = new Component("module", code, Codepage.Unknown, {
		sourceEncoding: "utf-8-bom"
	});

	await writeComponent(path, component);

	const writeCalls = (writeFile as jest.Mock).mock.calls;
	const writeCall = writeCalls[writeCalls.length - 1];
	const bytes = writeCall?.[1];
	if (!Buffer.isBuffer(bytes))
		throw new Error("Expected UTF-8 BOM output to be written as a Buffer");

	expect(bytes.subarray(0, 3)).toEqual(Buffer.from([0xef, 0xbb, 0xbf]));
	expect(bytes.subarray(3).toString("utf8")).toBe(code);
});
