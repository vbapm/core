import { wildcard } from "../../../tests/__fixtures__";
import { reset, setup } from "../../../tests/__helpers__/project";
import { resolveSourceFiles, resolveTargetPaths, classifyByPath, ResolvedSource } from "../apply-extract";
import { Component, ComponentType } from "../component";
import { Codepage } from "../encoding-sniffer";

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

describe("resolveTargetPaths", () => {
	function makeComponent(type: ComponentType, name: string): Component {
		const ext = type === "module" ? ".bas" : type === "form" ? ".frm" : ".cls";
		const code =
			type === "form"
				? `VERSION 5.00\r\nBegin Form\r\nEnd\r\nAttribute VB_Name = "${name}"\r\n`
				: `Attribute VB_Name = "${name}"\r\n`;
		return new Component(type, Buffer.from(code), Codepage.Unknown, {});
	}

	test("resolves target paths using folder and subfolders from src-properties", async () => {
		const { project } = await setup(wildcard);

		const components = [
			makeComponent("module", "NewModule"),
			makeComponent("class", "NewClass"),
			makeComponent("form", "NewForm"),
			makeComponent("object", "Sheet4")
		];

		const map = await resolveTargetPaths(project, components);

		expect(map.size).toBe(4);

		const paths = [...map.keys()];
		const pathStr = paths.join("|");
		expect(pathStr).toContain("src/Modules/NewModule.bas");
		expect(pathStr).toContain("src/Class Modules/NewClass.cls");
		expect(pathStr).toContain("src/Forms/NewForm.frm");
		expect(pathStr).toContain("src/Excel Objects/Sheet4.cls");
	});

	test("uses default src folder when folder is not set", async () => {
		const { project } = await setup(wildcard);
		project.manifest.srcProperties = undefined;

		const components = [makeComponent("module", "MyModule")];

		const map = await resolveTargetPaths(project, components);

		const [path] = [...map.keys()];
		expect(path).toEqual(expect.stringContaining("src/MyModule.bas"));
		expect(path).not.toEqual(expect.stringContaining("src/Modules/"));
	});
});

describe("classifyByPath", () => {
	function makeComponent(
		type: ComponentType,
		name: string,
		code?: string
	): Component {
		const body =
			code ??
			(type === "form"
				? `VERSION 5.00\r\nBegin Form\r\nEnd\r\nAttribute VB_Name = "${name}"\r\n`
				: `Attribute VB_Name = "${name}"\r\n`);
		return new Component(type, Buffer.from(body), Codepage.Unknown, {});
	}

	function makeResolved(name: string, type: ComponentType, code?: string): ResolvedSource {
		const component = makeComponent(type, name, code);
		return {
			source: { name, path: component.details.path! },
			component
		};
	}

	const PROJ = "/project";

	test("unchanged files are not listed", () => {
		const sources = new Map<string, ResolvedSource>();
		const targets = new Map<string, Component>();

		const src = makeResolved("Module1", "module", "Attribute VB_Name = \"Module1\"\r\n");
		src.component.details.path = PROJ + "/src/Modules/Module1.bas";
		src.source.path = PROJ + "/src/Modules/Module1.bas";
		sources.set(PROJ + "/src/Modules/Module1.bas", src);

		const tgt = makeComponent("module", "Module1", "Attribute VB_Name = \"Module1\"\r\n");
		targets.set(PROJ + "/src/Modules/Module1.bas", tgt);

		const result = classifyByPath(sources, targets);

		expect(result.modified).toHaveLength(0);
		expect(result.created).toHaveLength(0);
		expect(result.orphaned).toHaveLength(0);
	});

	test("detects modified files (same path, different code)", () => {
		const sources = new Map<string, ResolvedSource>();
		const targets = new Map<string, Component>();

		const src = makeResolved("Module1", "module", 'Attribute VB_Name = "Module1"\r\nold code\r\n');
		src.component.details.path = PROJ + "/src/Modules/Module1.bas";
		src.source.path = PROJ + "/src/Modules/Module1.bas";
		sources.set(PROJ + "/src/Modules/Module1.bas", src);

		const tgt = makeComponent("module", "Module1", 'Attribute VB_Name = "Module1"\r\nnew code\r\n');
		targets.set(PROJ + "/src/Modules/Module1.bas", tgt);

		const result = classifyByPath(sources, targets);

		expect(result.modified).toHaveLength(1);
		expect(result.modified[0].component.name).toBe("Module1");
		expect(result.created).toHaveLength(0);
		expect(result.orphaned).toHaveLength(0);
	});

	test("detects created files (only in targets)", () => {
		const sources = new Map<string, ResolvedSource>();
		const targets = new Map<string, Component>();

		const tgt = makeComponent("module", "NewModule");
		targets.set(PROJ + "/src/Modules/NewModule.bas", tgt);

		const result = classifyByPath(sources, targets);

		expect(result.created).toHaveLength(1);
		expect(result.created[0].component.name).toBe("NewModule");
		expect(result.modified).toHaveLength(0);
		expect(result.orphaned).toHaveLength(0);
	});

	test("detects orphaned files (only in sources)", () => {
		const sources = new Map<string, ResolvedSource>();
		const targets = new Map<string, Component>();

		const src = makeResolved("OldModule", "module");
		src.component.details.path = PROJ + "/src/Modules/OldModule.bas";
		src.source.path = PROJ + "/src/Modules/OldModule.bas";
		sources.set(PROJ + "/src/Modules/OldModule.bas", src);

		const result = classifyByPath(sources, targets);

		expect(result.orphaned).toHaveLength(1);
		expect(result.orphaned[0].component.name).toBe("OldModule");
		expect(result.orphaned[0].source?.name).toBe("OldModule");
		expect(result.modified).toHaveLength(0);
		expect(result.created).toHaveLength(0);
	});
});
