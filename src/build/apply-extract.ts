import walk from "walk-sync";
import { yellowBright } from "@timhall/ansi-colors";
import { env } from "../env";
import { Message } from "../messages";
import { resolveSrcFolder, resolveSrcSubfolders, writeManifest } from "../manifest";
import { Reference } from "../manifest/reference";
import { Source } from "../manifest/source";
import { Project } from "../project";
import { remove } from "../utils/fs";
import { parallel } from "../utils/parallel";
import { join, relative } from "../utils/path";
import { Codepage, labelToCodepage } from "./encoding-sniffer";
import { Component } from "./component";
import { isCoveredByWildcard, writeComponent } from "./apply-changeset";

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
export async function resolveSourceFiles(project: Project): Promise<Map<string, ResolvedSource>> {
	const map = new Map<string, ResolvedSource>();

	for (const source of project.manifest.src) {
		// Resolve encoding: per-source override > project-level > Unknown
		const declaredLabel = source.encoding ?? project.manifest.srcProperties?.encoding;
		const codepage = declaredLabel ? labelToCodepage(declaredLabel) : Codepage.Unknown;

		if (source.path.includes("*")) {
			// Wildcard — expand against the project directory.
			// Normalise to forward slashes for startsWith — source.path uses
			// / but project.paths.dir may use \ on Windows.
			const baseDir = project.paths.dir;
			const baseDirFwd = baseDir.replace(/\\/g, "/");
			const pattern = source.path.startsWith(baseDirFwd)
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
			const component = await Component.load(source.path, codepage, { binary_path: source.binary });
			component.details.sourceEncoding = declaredLabel;

			// Detect name mismatch: the manifest key differs from the
			// component's actual Attribute VB_Name in the file.
			if (source.name !== component.name && !source.path.includes("*")) {
				env.reporter.log(
					Message.SourceNameMismatch,
					yellowBright(
						`"${source.name}" in [source.files] was renamed to "${component.name}" ` +
							`(file has Attribute VB_Name = "${component.name}")`
					)
				);
			}

			map.set(source.path, { source, component });
		}
	}

	return map;
}

/**
 * Resolve target file paths for exported components based on the project's
 * `[source]` (folder + subfolders).
 *
 * Each component is placed under `folder/subfolder/component.filename`, where:
 * - `folder` comes from `[source].folder` (default `"src"`)
 * - `subfolder` maps the component type via `[source].subfolders`
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
	/** Entries whose manifest name differs from the actual component name. */
	renamed: { source: Source; newName: string }[];
	references: {
		added: Reference[];
		removed: Reference[];
	};
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
		orphaned: [],
		renamed: [],
		references: { added: [], removed: [] }
	};

	for (const [path, component] of targets) {
		const resolved = sources.get(path);
		if (resolved) {
			// Same path — check if code changed
			if (component.code !== resolved.component.code) {
				result.modified.push({ component });
			}
			// Detect name mismatches for individual entries (not wildcards):
			// - Source name ≠ file Attribute: user renamed manifest key
			// - Export name ≠ source name: workbook component was renamed
			if (!resolved.source.path.includes("*")) {
				if (resolved.source.name !== resolved.component.name) {
					// Manifest key differs from file content
					result.renamed.push({
						source: resolved.source,
						newName: resolved.component.name
					});
				} else if (component.name !== resolved.source.name) {
					// Export component was renamed; manifest key should follow
					result.renamed.push({
						source: resolved.source,
						newName: component.name
					});
					env.reporter.log(
						Message.SourceNameMismatch,
						yellowBright(
							`"${resolved.source.name}" in [source.files] was renamed to "${component.name}" ` +
								`(component in Excel file has Attribute VB_Name = "${component.name}")`
						)
					);
				}
			}
			sources.delete(path);
		} else {
			// New file
			result.created.push({ component });
		}
	}

	// Remaining sources are orphaned (no matching target)
	for (const resolved of sources.values()) {
		result.orphaned.push({
			component: resolved.component,
			source: resolved.source
		});
	}

	return result;
}

// ---------------------------------------------------------------------------
// Execution
// ---------------------------------------------------------------------------

/**
 * Execute a classified extract: write modified/created files, delete orphaned
 * files, re-scan wildcards to decide which created files need individual
 * `[src]` entries, and update the manifest.
 */
export async function applyExtract(project: Project, classified: ClassifiedExtract): Promise<void> {
	// --- Write modified files ---
	await parallel(classified.modified, item =>
		writeComponent(item.component.details.path!, item.component)
	);

	// --- Write created files ---
	await parallel(classified.created, item =>
		writeComponent(item.component.details.path!, item.component)
	);

	// --- Delete orphaned files ---
	await parallel(classified.orphaned, async item => {
		await remove(item.component.details.path!);
		// Also remove binary companion (.frx) if present
		if (item.component.binaryPath) {
			const dir = item.component.details.path!.replace(/[/\\][^/\\]*$/, "");
			await remove(join(dir, item.component.binaryPath));
		}
	});

	// --- Re-scan wildcards for coverage ---
	const needsEntry: ClassifiedComponent[] = [];
	for (const item of classified.created) {
		const relPath = relative(project.paths.dir, item.component.details.path!);
		if (!isCoveredByWildcard(relPath, project.manifest.src, project.paths.dir)) {
			needsEntry.push(item);
		}
	}

	// --- Update manifest ---
	const hasChanges =
		needsEntry.length > 0 ||
		classified.orphaned.length > 0 ||
		classified.renamed.length > 0 ||
		classified.references.added.length > 0 ||
		classified.references.removed.length > 0;

	if (hasChanges) {
		updateManifestForExtract(
			project,
			needsEntry,
			classified.orphaned,
			classified.renamed,
			classified.references
		);

		await writeManifest(project.manifest, project.paths.dir);
	}
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Update the project manifest: add individual `[src]` entries for created
 * files that aren't covered by wildcards, and remove entries for orphaned files.
 */
function updateManifestForExtract(
	project: Project,
	needsEntry: ClassifiedComponent[],
	orphaned: ClassifiedComponent[],
	renamed: { source: Source; newName: string }[],
	references: { added: Reference[]; removed: Reference[] }
): void {
	const src = project.manifest.src;

	// Rename entries whose manifest key differs from the actual component name
	for (const { source, newName } of renamed) {
		const index = src.findIndex((s: Source) => s.name === source.name && s.path === source.path);
		if (index >= 0) {
			src[index] = { ...source, name: newName };
		}
	}

	// Add individual entries for uncovered created files
	for (const item of needsEntry) {
		const sub = resolveSrcSubfolders(
			project.manifest.srcProperties?.subfolders,
			item.component.type
		);
		const folder = resolveSrcFolder(project.manifest.srcProperties);
		const srcPath = sub
			? `${folder}/${sub}/${item.component.filename}`
			: `${folder}/${item.component.filename}`;
		const entry: Source = {
			name: item.component.name,
			path: join(project.paths.dir, srcPath)
		};

		// Replace an existing entry with the same name if it points
		// to a different file (e.g. Module1 was "src/Validation.bas"
		// but the workbook now has a real Module1 component).
		const existingIndex = src.findIndex((s: Source) => s.name === item.component.name);
		if (existingIndex >= 0) {
			src[existingIndex] = entry;
		} else {
			src.push(entry);
		}
	}

	// Remove entries for orphaned files
	for (const item of orphaned) {
		const index = src.findIndex((s: Source) => s.name === item.component.name);
		if (index >= 0) src.splice(index, 1);
	}

	// Add new references
	for (const ref of references.added) {
		project.manifest.references.push(ref);
	}

	// Remove orphaned references
	for (const ref of references.removed) {
		const index = project.manifest.references.findIndex((r: Reference) => r.name === ref.name);
		if (index >= 0) project.manifest.references.splice(index, 1);
	}
}
