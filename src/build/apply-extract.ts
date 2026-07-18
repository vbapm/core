import walk from "walk-sync";
import { resolveSrcFolder, resolveSrcSubfolders } from "../manifest";
import { Source } from "../manifest/source";
import { Project } from "../project";
import { join, relative } from "../utils/path";
import { Codepage, labelToCodepage } from "./encoding-sniffer";
import { Component } from "./component";

export interface ResolvedSource {
	source: Source;
	component: Component;
}

/**
 * Resolve the project manifest's `[src]` entries to a concrete map of
 * file paths → components.
 *
 * Wildcard entries (containing `*`) are expanded via `walk-sync` against the
 * project directory. Single-path entries are loaded directly. Only the project
 * manifest's own `[src]` is resolved — dependencies and `[dev-src]` are excluded.
 *
 * @returns Map keyed by absolute file path.
 */
export async function resolveSourceFiles(
	project: Project
): Promise<Map<string, ResolvedSource>> {
	const map = new Map<string, ResolvedSource>();

	for (const source of project.manifest.src) {
		// Resolve encoding: per-source override > project-level > Unknown
		const declaredLabel = source.encoding ?? project.manifest.srcProperties?.encoding;
		const codepage = declaredLabel ? labelToCodepage(declaredLabel) : Codepage.Unknown;

		if (source.path.includes("*")) {
			// Wildcard — expand against the project directory
			const baseDir = project.paths.dir;
			const pattern = source.path.startsWith(baseDir)
				? relative(baseDir, source.path)
				: source.path;
			const matched = walk(baseDir, { globs: [pattern], directories: false });
			for (const file of matched) {
				const absPath = join(baseDir, file);
				const component = await Component.load(absPath, codepage);
				component.details.sourceEncoding = declaredLabel;
				map.set(absPath, { source, component });
			}
		} else {
			// Single path — load directly
			const component = await Component.load(
				source.path,
				codepage,
				{ binary_path: source.binary }
			);
			component.details.sourceEncoding = declaredLabel;
			map.set(source.path, { source, component });
		}
	}

	return map;
}

/**
 * Resolve target file paths for exported components based on the project's
 * `[src-properties]` (folder + subfolders).
 *
 * Each component is placed under `folder/subfolder/component.filename`, where:
 * - `folder` comes from `[src-properties].folder` (default `"src"`)
 * - `subfolder` maps the component type via `[src-properties].subfolders`
 *
 * @returns Map keyed by absolute target file path.
 */
export function resolveTargetPaths(
	project: Project,
	components: Component[]
): Map<string, Component> {
	const map = new Map<string, Component>();
	const folder = resolveSrcFolder(project.manifest.srcProperties);
	const subfolders = project.manifest.srcProperties?.subfolders;

	for (const component of components) {
		const sub = resolveSrcSubfolders(subfolders, component.type);
		const path = join(project.paths.dir, folder, sub, component.filename);
		component.details.path = path;
		map.set(path, component);
	}

	return map;
}

// ---------------------------------------------------------------------------
// Classification
// ---------------------------------------------------------------------------

export interface ClassifiedComponent {
	component: Component;
	/** Present when the file is orphaned — records the original Source entry. */
	source?: Source;
}

export interface ClassifiedExtract {
	modified: ClassifiedComponent[];
	created: ClassifiedComponent[];
	orphaned: ClassifiedComponent[];
}

/**
 * Classify components by comparing the on-disk source map with the resolved
 * target paths from the export.
 *
 * - Path in both maps, same code → unchanged (not included)
 * - Path in both maps, different code → **modified**
 * - Path only in targets → **created**
 * - Path only in sources → **orphaned**
 */
export function classifyByPath(
	sources: Map<string, ResolvedSource>,
	targets: Map<string, Component>
): ClassifiedExtract {
	const result: ClassifiedExtract = {
		modified: [],
		created: [],
		orphaned: []
	};

	for (const [path, component] of targets) {
		const resolved = sources.get(path);
		if (resolved) {
			// Same path — check if code changed
			if (component.code !== resolved.component.code) {
				result.modified.push({ component });
			}
			sources.delete(path);
		} else {
			// New file
			result.created.push({ component });
		}
	}

	// Remaining sources are orphaned (no matching target)
	for (const [path, resolved] of sources) {
		result.orphaned.push({
			component: resolved.component,
			source: resolved.source
		});
	}

	return result;
}
