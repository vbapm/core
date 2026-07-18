import walk from "walk-sync";
import { Source } from "../manifest/source";
import { Project } from "../project";
import { join, relative } from "../utils/path";
import { Codepage, labelToCodepage } from "./encoding-sniffer";
import { Component, extensionToType } from "./component";

export interface ResolvedSource {
	source: Source;
	component: Component;
}

/**
 * Phase 1: Resolve the project manifest's `[src]` entries to a concrete map of
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
