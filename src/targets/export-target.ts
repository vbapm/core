import dedent from "@timhall/dedent";
import {
	applyExtract,
	classifyByPath,
	loadFromExport,
	resolveSourceFiles,
	resolveTargetPaths,
	toSrc
} from "../build";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { Target } from "../manifest/target";
import { Project } from "../project";
import { copy, ensureDir, pathExists, remove } from "../utils/fs";
import { dirname, join } from "../utils/path";
import { unzip } from "../utils/zip";
import { ProjectInfo } from "./project-info";
import { filterTarget, mapTarget } from "./transform-target";
import { normalizeWorksheetNames } from "./transforms/normalize-worksheet-names";

export interface ExportOptions {
	xmlOnly?: boolean;
	vbaOnly?: boolean;
	skipSheetNameNormalization?: boolean;
}

/**
 * Export target (with staging directory)
 *
 * 1. Export source from target to staging (done previously)
 * 2. Extract target to staging
 * 3. Export build graph to src
 * 4. Move extracted to target to src
 */
export async function exportTarget(
	target: Target,
	info: ProjectInfo,
	staging: string,
	options: ExportOptions = {}
) {
	const { project, dependencies, blankTarget } = info;
	const { xmlOnly = false, vbaOnly = false, skipSheetNameNormalization = false } = options;

	// Extract target to staging
	let extracted: string;
	if (!blankTarget && !vbaOnly) {
		extracted = await extractTarget(project, target, staging);
		if (!skipSheetNameNormalization) {
			await normalizeWorksheetNames(extracted);
		}
	}

	if (!xmlOnly) {
		// Compare project and exported and apply changes to project
		const sources = await resolveSourceFiles(project);
		const exported_build_graph = await loadFromExport(staging);
		const transformed_build_graph = await toSrc(exported_build_graph);

		// Exclude dependency-owned components from the export so they
		// aren't treated as project files
		const depNames = new Set(dependencies.flatMap(m => m.src.map(s => s.name)));
		const projectComponents = transformed_build_graph.components.filter(c => !depNames.has(c.name));

		const targets = resolveTargetPaths(project, projectComponents);
		const classified = classifyByPath(new Map(sources), targets);

		// Compare references by name (unchanged from original compareBuildGraphs logic)
		const existingRefs = new Map(project.manifest.references.map(r => [r.name, r]));
		for (const ref of transformed_build_graph.references) {
			const existing = existingRefs.get(ref.name);
			if (existing) {
				existingRefs.delete(ref.name);
			} else {
				classified.references.added.push(ref);
			}
		}
		// Remaining in existingRefs are no longer in the workbook
		for (const ref of existingRefs.values()) {
			classified.references.removed.push(ref);
		}

		await applyExtract(project, classified);
	}

	// Move target to dest
	if (!blankTarget && !vbaOnly) {
		await remove(target.path);
		await copy(extracted!, target.path);
	}

	// Finally, cleanup staging
	await remove(staging);
}

export async function extractTarget(
	project: Project,
	target: Target,
	staging: string
): Promise<string> {
	let src = join(project.paths.build, target.filename);
	const dest = join(staging, "target");

	if (!(await pathExists(src))) {
		throw new CliError(
			ErrorCode.ExportTargetNotFound,
			dedent`
        Could not find built target for type "${target.type}"
        (checked "${src}").
      `
		);
	}

	// For Mac, stage target to avoid permission prompts
	if (!env.isWindows) {
		const staged = join(staging, "staged", target.filename);
		if (!(await pathExists(staged))) {
			await ensureDir(dirname(staged));
			await copy(src, staged);
		}

		src = staged;
	}

	await ensureDir(dest);
	await unzip(src, dest, { filter: filterTarget, map: mapTarget });

	return dest;
}
