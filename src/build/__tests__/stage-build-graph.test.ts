import { complex, dir, standardImport } from "../../../tests/__fixtures__";
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

	const component = new Component(
		"module",
		`Attribute VB_Name = "AsciiModule"\n' Café ${String.fromCodePoint(0x2242, 0x338)} ${String.fromCodePoint(0xac00)} ${String.fromCodePoint(0x1f600)}\n`,
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
	} finally {
		rmSync(stagingDir, { recursive: true, force: true });
	}
});
