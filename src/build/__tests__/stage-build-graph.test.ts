import { complex, dir, empty, standardImport } from "../../../tests/__fixtures__";
import { reset, setup } from "../../../tests/__helpers__/project";
import { pathExists, readFile, ensureDir } from "../../utils/fs";
import { join, relative } from "../../utils/path";
import { writeFileSync, mkdirSync, rmSync } from "fs";
import iconv from "iconv-lite";
import { ImportGraph } from "../build-graph";
import { Component } from "../component";
import { Codepage } from "../encoding-sniffer";
import { loadFromProject } from "../load-from-project";
import { stageBuildGraph } from "../stage-build-graph";
import { toSrc } from "../transform-build-graph";

afterEach(reset);

test("should stage BuildGraph", async () => {
	expect.assertions(15);

	const { project, dependencies } = await setup(complex);
	const graph = await loadFromProject(project, dependencies);

	const import_graph = await stageBuildGraph(graph, standardImport);
	expect(normalizeImportGraph(import_graph)).toMatchSnapshot();

	for (const source of import_graph.components) {
		expect(await pathExists(source.path)).toEqual(true);
	}
});

test("stages Power Query formulas as UTF-8 files", async () => {
	const { project, dependencies } = await setup(empty);
	const graph = await loadFromProject(project, dependencies);
	graph.powerQueries = [{ name: "Query", formula: 'let Source = "café" in Source' }];

	const importGraph = await stageBuildGraph(graph, standardImport);
	const stagedQuery = importGraph.powerQueries?.[0];

	expect(stagedQuery).toEqual({ name: "Query", path: expect.any(String) });
	expect(await pathExists(stagedQuery!.path)).toBe(true);
	expect(await readFile(stagedQuery!.path, "utf8")).toBe(graph.powerQueries[0].formula);
});

function normalizeImportGraph(
	graph: ImportGraph
): Omit<ImportGraph, "components"> & { components: { name: string; path: string }[] } {
	const { name, references } = graph;
	const components = graph.components.map(source => {
		return {
			name: source.name,
			path: relative(dir, source.path)
		};
	});

	return { name, components, references };
}

test("transcodes when source encoding differs from target", async () => {
	const fixtureDir = join(dir, "projects", "transcode-test");

	mkdirSync(join(fixtureDir, "src"), { recursive: true });
	mkdirSync(join(fixtureDir, "targets"), { recursive: true });

	const toml = `\
[package]
name = "transcode-test"
version = "1.0.0"
authors = ["Test"]

[source]
encoding = "cp1252"

[src]
Module1 = "src/Module1.bas"
`;

	// CP1252: é = 0xE9, à = 0xE0
	const srcContent =
		'Attribute VB_Name = "Module1"\n\' Déjà vu – naïve\nPublic Sub Bonjour()\nEnd Sub\n';
	const srcPath = join(fixtureDir, "src", "Module1.bas");
	const manifestPath = join(fixtureDir, "vbaproject.toml");
	writeFileSync(srcPath, srcContent);
	writeFileSync(manifestPath, toml);

	const { project, dependencies } = await setup(fixtureDir);
	const graph = await loadFromProject(project, dependencies);

	// Stage with target CP932 (Japanese) — different from source CP1252
	const stagingDir = join(project.paths.staging, "transcode-import");
	await ensureDir(stagingDir);
	const import_graph = await stageBuildGraph(graph, stagingDir, Codepage.Windows932);

	const stagedPath = import_graph.components[0].path;
	expect(await pathExists(stagedPath)).toBe(true);

	// Read as binary — should be CP932-encoded, not CP1252
	const buf = await readFile(stagedPath);
	expect(buf).toBeDefined();

	// Decode as CP932 to verify the content survived transcoding
	const decoded = iconv.decode(buf, "cp932");
	expect(decoded).toContain("Bonjour");

	// Cleanup
	rmSync(fixtureDir, { recursive: true, force: true });
});

test("uses named HTML entities, then hex references for unmapped characters in ASCII targets", async () => {
	const stagingDir = join(dir, "projects", "ascii-encoding-stage");
	await ensureDir(stagingDir);

	const originalComment = `Café ${String.fromCodePoint(0x2242, 0x338)} ${String.fromCodePoint(0xac00)} ${String.fromCodePoint(0x1f600)}`;
	const component = new Component(
		"module",
		`Attribute VB_Name = "AsciiModule"\n' ${originalComment}\nPublic Const EntityText As String = "&eacute;"\n`,
		Codepage.Unknown,
		{ sourceEncoding: "utf-8" }
	);
	const graph = {
		name: "ascii-encoding-test",
		components: [component],
		references: [],
		fromDependencies: { components: new Map(), references: new Map() }
	};

	try {
		const importGraph = await stageBuildGraph(graph, stagingDir, Codepage.Ascii);
		const output = await readFile(importGraph.components[0].path);
		const encodedText = output.toString("ascii");

		expect(encodedText).toContain("Caf&eacute;");
		expect(encodedText).toContain("&nesim;");
		expect(encodedText).toContain("&#xac00;");
		expect(encodedText).toContain("&#x1f600;");
		expect(output.every(byte => byte <= 0x7f)).toBe(true);

		const extracted = await toSrc({
			...graph,
			components: [new Component("module", output, Codepage.Ascii)]
		});
		expect(extracted.components[0].code).toContain(`' ${originalComment}`);
		expect(extracted.components[0].code).toContain(
			'Public Const EntityText As String = "&eacute;"'
		);
	} finally {
		rmSync(stagingDir, { recursive: true, force: true });
	}
});

test("wraps escaped comments within VBA's physical-line limit and reassembles them on extraction", async () => {
	const stagingDir = join(dir, "projects", "long-comment-stage");
	await ensureDir(stagingDir);

	const originalComment = "&eacute;".repeat(86);
	const originalCode = `Attribute VB_Name = "LongComment"\n' ${originalComment}\n`;
	const component = new Component("module", originalCode, Codepage.Unknown);
	const graph = {
		name: "long-comment-test",
		components: [component],
		references: [],
		fromDependencies: { components: new Map(), references: new Map() }
	};

	try {
		const importGraph = await stageBuildGraph(graph, stagingDir, Codepage.Ascii);
		const output = await readFile(importGraph.components[0].path);
		const encodedText = output.toString("ascii");
		const commentLines = encodedText.split(/\r?\n/).filter(line => line.startsWith("'"));

		expect(commentLines).toHaveLength(2);
		expect(commentLines.every(line => line.length <= 1023)).toBe(true);
		expect(commentLines[0]).toMatch(/ _$/);

		const extracted = await toSrc({
			...graph,
			components: [new Component("module", output, Codepage.Ascii)]
		});
		expect(extracted.components[0].code).toContain(`' ${originalComment}`);
	} finally {
		rmSync(stagingDir, { recursive: true, force: true });
	}
});

test("rejects comments that exceed VBA's line-continuation limit", async () => {
	const stagingDir = join(dir, "projects", "oversized-comment-stage");
	await ensureDir(stagingDir);

	const component = new Component(
		"module",
		`Attribute VB_Name = "OversizedComment"\n' ${"x".repeat(26_000)}\n`,
		Codepage.Unknown
	);
	const graph = {
		name: "oversized-comment-test",
		components: [component],
		references: [],
		fromDependencies: { components: new Map(), references: new Map() }
	};

	try {
		await expect(stageBuildGraph(graph, stagingDir, Codepage.Ascii)).rejects.toThrow(
			"maximum of 24 line continuations"
		);
	} finally {
		rmSync(stagingDir, { recursive: true, force: true });
	}
});
