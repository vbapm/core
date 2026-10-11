import dedent from "@timhall/dedent";
import { yellowBright } from "@timhall/ansi-colors";
import walk from "walk-sync";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { Message } from "../messages";
import { Manifest } from "../manifest";
import { Reference } from "../manifest/reference";
import { Project } from "../project";
import { BuildOptions } from "../targets/build-target";
import { readFile } from "../utils/fs";
import { join, relative } from "../utils/path";
import { joinCommas } from "../utils/text";
import { BuildGraph, FromDependences } from "./build-graph";
import { Codepage, labelToCodepage, SUPPORTED_WINDOWS_CODEPAGE_LABELS } from "./encoding-sniffer";
import { byComponentTypeThenName, Component } from "./component";

import jschardet from "jschardet";

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
		for (const source of manifest.src) {
			// Resolve encoding: per-source override > project-level > Unknown
			const declaredLabel = source.encoding ?? manifest.srcProperties?.encoding;
			const codepage = declaredLabel ? labelToCodepage(declaredLabel) : Codepage.Unknown;

			// Expand wildcards against the project directory
			if (source.path.includes("*")) {
				const baseDir = project.paths.dir;
				const normalizedSourcePath = source.path.replace(/\\/g, "/");
				const normalizedBaseDir = baseDir.replace(/\\/g, "/");
				const pattern = normalizedSourcePath.startsWith(normalizedBaseDir)
					? relative(baseDir, source.path)
					: source.path;
				const matched = walk(baseDir, { globs: [pattern], directories: false });
				for (const file of matched) {
					const absPath = join(baseDir, file);
					loadingComponents.push(
						Component.load(absPath, codepage, {}).then(component => {
							component.details.sourceEncoding = declaredLabel;
							if (manifest !== project.manifest) {
								fromDependencies.components.set(component, manifest.name);
							}
							return component;
						})
					);
				}
			} else {
				loadingComponents.push(
					Component.load(source.path, codepage, { binary_path: source.binary }).then(component => {
						component.details.sourceEncoding = declaredLabel;
						if (manifest !== project.manifest) {
							fromDependencies.components.set(component, manifest.name);
						}
						return component;
					})
				);
			}
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
		for (const source of project.manifest.devSrc) {
			if (source.path.includes("*")) {
				const baseDir = project.paths.dir;
				const normalizedSourcePath = source.path.replace(/\\/g, "/");
				const normalizedBaseDir = baseDir.replace(/\\/g, "/");
				const pattern = normalizedSourcePath.startsWith(normalizedBaseDir)
					? relative(baseDir, source.path)
					: source.path;
				const matched = walk(baseDir, { globs: [pattern], directories: false });
				for (const file of matched) {
					const absPath = join(baseDir, file);
					loadingComponents.push(
						Component.load(absPath, Codepage.Unknown, { binary_path: source.binary })
					);
				}
			} else {
				loadingComponents.push(
					Component.load(source.path, Codepage.Unknown, { binary_path: source.binary })
				);
			}
		}
		for (const reference of project.manifest.devReferences) {
			const nameGuid = `${reference.name}_${reference.guid}`;
			if (foundReferences[nameGuid]) continue;

			references.push(reference);
			foundReferences[nameGuid] = true;
		}
	}

	const components = (await Promise.all(loadingComponents)).sort(byComponentTypeThenName);

	// Validate that non-wildcard [source.files] keys match their file's Attribute VB_Name
	validateSrcNames(project, components);

	const graph = {
		name: project.manifest.codename || "VBAProject",
		components,
		references,
		fromDependencies
	};

	await validateEncoding(project, graph);
	validateGraph(project, graph);
	return graph;
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
	const srcEncoding = project.manifest.srcProperties?.encoding;

	for (const component of graph.components) {
		// Only check project-owned components, not dependency components
		if (graph.fromDependencies.components.has(component)) continue;

		// Check if encoding is declared (project-level or per-source)
		// VBA component names are case-insensitive
		const source = project.manifest.src.find(
			s => s.name.localeCompare(component.name, undefined, { sensitivity: "base" }) === 0
		);
		const declaredEncoding = source?.encoding ?? srcEncoding;
		if (declaredEncoding) continue;

		// Check for non-ASCII characters
		if (!hasNonAscii(component.code)) continue;

		// Try jschardet to suggest an encoding
		let suggestion = "";
		const sourcePath = source?.path;
		if (sourcePath) {
			try {
				const buffer = await readFile(sourcePath);

				const results = (
					jschardet.detectAll(buffer) as Array<{ encoding: string; confidence: number }>
				)
					.filter(r => SUPPORTED_WINDOWS_CODEPAGE_LABELS.has(r.encoding))
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
        Non-ASCII characters detected in "${source?.path || component.filename}".
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
