import dedent from "@timhall/dedent";
import { yellowBright } from "@timhall/ansi-colors";
import walk from "walk-sync";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { Message } from "../messages";
import { Manifest, resolveSrcFolder, resolveSrcSubfolders } from "../manifest";
import { Reference } from "../manifest/reference";
import {
	isPowerQueryPath,
	isPowerQuerySource,
	resolveSourceEncoding,
	Source,
	sourcePathMatches
} from "../manifest/source";
import { Project } from "../project";
import { BuildOptions } from "../targets/build-target";
import { readFile } from "../utils/fs";
import { join, relative } from "../utils/path";
import { joinCommas } from "../utils/text";
import { BuildGraph, FromDependences } from "./build-graph";
import { loadPowerQueries } from "./load-power-queries";
import { Codepage, labelToCodepage, SUPPORTED_WINDOWS_CODEPAGE_LABELS } from "./encoding-sniffer";
import { byComponentTypeThenName, Component } from "./component";

import jschardet from "jschardet";

/**
 * Loads project and dependency sources, references and Power Query files into
 * a build graph, then validates the graph and its source encodings.
 *
 * Development sources and dependencies are included unless release mode is
 * enabled.
 *
 * @param project - Project whose manifest and source files will be loaded.
 * @param dependencies - Resolved dependency manifests to include in the graph.
 * @param options - Build options controlling release filtering.
 * @returns The validated build graph.
 * @throws {CliError} If graph validation fails or a source encoding is required.
 */
export async function loadFromProject(
	project: Project,
	dependencies: Manifest[],
	options: BuildOptions = {}
): Promise<BuildGraph> {
	let includedDependencies = dependencies;
	if (options.release) {
		const devDependencies = project.manifest.devDependencies.map(dependency => dependency.name);
		includedDependencies = dependencies.filter(manifest => {
			return !devDependencies.includes(manifest.name);
		});
	}

	const manifests = [project.manifest, ...includedDependencies];
	const loadingComponents: Promise<Component>[] = [];
	const references: Reference[] = [];
	const foundReferences: { [name_guid: string]: boolean } = {};
	const fromDependencies: FromDependences = { components: new Map(), references: new Map() };

	// Load components and references from project and dependencies
	for (const manifest of manifests) {
		for (const [sourcePath, source] of expandManifestSources(manifest.src, project.paths.dir)) {
			const declaredLabel = resolveSourceEncoding(
				manifest.src,
				sourcePath,
				project.paths.dir,
				manifest.srcProperties?.encoding
			);
			const codepage = declaredLabel ? labelToCodepage(declaredLabel) : Codepage.Unknown;

			loadingComponents.push(
				Component.load(
					sourcePath,
					codepage,
					source.path.includes("*") ? {} : { binary_path: source.binary }
				).then(component => {
					markObjectComponent(component, project, sourcePath);
					component.details.sourceEncoding = declaredLabel;
					if (manifest !== project.manifest) {
						fromDependencies.components.set(component, manifest.name);
					}
					return component;
				})
			);
		}
		for (const reference of manifest.references) {
			const nameGuid = `${reference.name}_${reference.guid}`;
			if (foundReferences[nameGuid]) continue;

			references.push(reference);
			if (manifest !== project.manifest) {
				fromDependencies.references.set(reference, manifest.name);
			}

			foundReferences[nameGuid] = true;
		}
	}

	if (!options.release) {
		for (const [sourcePath, source] of expandManifestSources(
			project.manifest.devSrc,
			project.paths.dir
		)) {
			const declaredLabel = resolveSourceEncoding(
				project.manifest.devSrc,
				sourcePath,
				project.paths.dir,
				project.manifest.srcProperties?.encoding
			);
			const codepage = declaredLabel ? labelToCodepage(declaredLabel) : Codepage.Unknown;
			loadingComponents.push(
				Component.load(
					sourcePath,
					codepage,
					source.path.includes("*") ? {} : { binary_path: source.binary }
				).then(component => {
					component.details.sourceEncoding = declaredLabel;
					return component;
				})
			);
		}
		for (const reference of project.manifest.devReferences) {
			const nameGuid = `${reference.name}_${reference.guid}`;
			if (foundReferences[nameGuid]) continue;

			references.push(reference);
			foundReferences[nameGuid] = true;
		}
	}

	const components = (await Promise.all(loadingComponents)).sort(byComponentTypeThenName);
	const powerQueries = await loadPowerQueries(project, options);

	// Validate that non-wildcard [source.files] keys match their file's Attribute VB_Name
	validateSrcNames(project, components);

	const graph = {
		name: project.manifest.codename || "VBAProject",
		components,
		references,
		fromDependencies,
		...(powerQueries.length > 0 ? { powerQueries } : {})
	};

	await validateEncoding(project, graph);
	validateGraph(project, graph);
	return graph;
}

function expandManifestSources(sources: Source[], projectDir: string): Map<string, Source> {
	const paths = new Map<string, Source>();

	for (const source of sources) {
		if (isPowerQuerySource(source)) continue;

		if (!source.path.includes("*")) {
			if (!paths.has(source.path)) paths.set(source.path, source);
			continue;
		}

		const normalizedSourcePath = source.path.replace(/\\/g, "/");
		const normalizedBaseDir = projectDir.replace(/\\/g, "/");
		const pattern = normalizedSourcePath.startsWith(normalizedBaseDir)
			? relative(projectDir, source.path)
			: source.path;
		const matched = walk(projectDir, { globs: [pattern], directories: false });
		for (const file of matched) {
			if (isPowerQueryPath(file)) continue;

			const sourcePath = join(projectDir, file);
			const hasExactEntry = sources.some(
				entry => !entry.path.includes("*") && sourcePathMatches(entry.path, sourcePath, projectDir)
			);
			if (!hasExactEntry && !paths.has(sourcePath)) paths.set(sourcePath, source);
		}
	}

	return paths;
}

function validateGraph(project: Project, graph: BuildGraph) {
	const componentsByName: { [name: string]: string[] } = {};
	const referencesByName: { [name: string]: Reference[] } = {};
	const errors = [];

	for (const component of graph.components) {
		if (!componentsByName[component.name]) componentsByName[component.name] = [];

		const manifestName = graph.fromDependencies.components.get(component) || project.manifest.name;
		componentsByName[component.name].push(manifestName);
	}
	for (const reference of graph.references) {
		if (!referencesByName[reference.name]) referencesByName[reference.name] = [];
		referencesByName[reference.name].push(reference);
	}

	for (const [name, from] of Object.entries(componentsByName)) {
		if (from.length > 1) {
			const names = from.map(name => `"${name}"`);
			errors.push(`Source "${name}" is present in manifests named ${joinCommas(names)}`);
		}
	}
	for (const [name, references] of Object.entries(referencesByName)) {
		if (references.length > 1) {
			const versions = references.map(reference => `${reference.major}.${reference.minor}`);
			errors.push(`Reference "${name}" has multiple versions: ${joinCommas(versions)}`);
		}
	}

	if (errors.length) {
		throw new CliError(
			ErrorCode.BuildInvalid,
			dedent`
        Invalid build:

        ${errors.join("\n")}
      `
		);
	}
}

/**
 * Validate that any source file containing non-ASCII characters has
 * an encoding declared (encoding in [source] or encoding on
 * the individual source entry). If not, fail with a jschardet
 * suggestion.
 */
async function validateEncoding(project: Project, graph: BuildGraph) {
	for (const component of graph.components) {
		// Only check project-owned components, not dependency components
		if (graph.fromDependencies.components.has(component)) continue;

		if (component.details.sourceEncoding) continue;

		// Check for non-ASCII characters
		if (!hasNonAscii(component.code)) continue;

		const sourcePath = component.details.path;
		let suggestion = "";
		if (sourcePath) {
			try {
				const buffer = await readFile(sourcePath);

				const results = (
					jschardet.detectAll(buffer) as Array<{ encoding: string; confidence: number }>
				)
					.filter(r => SUPPORTED_WINDOWS_CODEPAGE_LABELS.has(r.encoding.toLowerCase()))
					.sort((a, b) => b.confidence - a.confidence);

				if (results.length > 0 && results[0].confidence >= 0.5) {
					const label = results[0].encoding.toLowerCase();
					suggestion =
						`\nSuggested change:\n\n  [source]\n  encoding = "${label}"` +
						`\n\n(Detection by jschardet, confidence: ${Math.round(results[0].confidence * 100)}%)`;
				}
			} catch (err) {
				suggestion =
					"\n\n(Unable to suggest an encoding." +
					(err instanceof Error ? ` ${err.message}` : "") +
					")";
			}
		}

		throw new CliError(
			ErrorCode.BuildInvalid,
			dedent`
        Non-ASCII characters detected in "${sourcePath || component.filename}".
        Please specify the encoding of the source code so the build process
        can preserve them correctly.${suggestion}
      `
		);
	}
}

/** Check if a string contains any non-ASCII characters (U+0080+). */
function hasNonAscii(str: string): boolean {
	for (let i = 0; i < str.length; i++) {
		if (str.charCodeAt(i) > 127) return true;
	}
	return false;
}

/**
 * Validate that every non-wildcard [source.files] entry's key matches the file's
 * Attribute VB_Name.  Mismatches could cause roundtrip renames
 * when exporting and re-importing, so we warn the user.
 */
function validateSrcNames(project: Project, components: Component[]): void {
	// Build a map from component path to component (for single-path loads)
	const compByPath = new Map<string, Component>();
	for (const c of components) {
		if (c.details.path) compByPath.set(c.details.path, c);
	}

	for (const source of project.manifest.src) {
		if (source.path.includes("*")) continue;

		const comp = compByPath.get(source.path);
		if (!comp) continue;

		// The loaded component's name comes from Attribute VB_Name.
		// Warn if the manifest key doesn't match.
		if (source.name !== comp.name) {
			env.reporter.log(
				Message.SourceNameMismatch,
				yellowBright(
					`WARN: "${source.name}" in [source.files] does not match Attribute VB_Name = "${comp.name}" ` +
						`in "${relative(project.paths.dir, source.path)}".\n` +
						`  Run "vbapm manifest fix" to rename the [source.files] key automatically.`
				)
			);
		}
	}
}

function markObjectComponent(component: Component, project: Project, sourcePath: string): void {
	const objectSubfolder = project.manifest.srcProperties?.subfolders?.Objects;
	if (!objectSubfolder) return;

	const sourceFolder = relative(project.paths.dir, sourcePath).toLowerCase();
	const objectFolder = join(
		resolveSrcFolder(project.manifest.srcProperties),
		resolveSrcSubfolders(project.manifest.srcProperties?.subfolders, "object")
	)
		.toLowerCase()
		.replace(/\/$/, "");

	if (sourceFolder.startsWith(`${objectFolder}/`)) component.type = "object";
}
