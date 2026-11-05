import { readFile } from "fs/promises";
import walk from "walk-sync";
import { basename, extname, join, relative } from "../utils/path";
import { Project } from "../project";
import { isPowerQueryPath, isPowerQuerySource, Source } from "../manifest/source";
import { PowerQuery } from "./build-graph";

export async function loadPowerQueries(
	project: Project,
	options: { release?: boolean } = {}
): Promise<PowerQuery[]> {
	const sources = [
		...project.manifest.src,
		...(options.release ? [] : project.manifest.devSrc)
	].filter(isPowerQuerySource);
	const queries: PowerQuery[] = [];
	const names = new Set<string>();

	for (const source of sources) {
		for (const file of resolvePowerQueryFiles(project, source)) {
			const name = source.path.includes("*") ? basename(file, extname(file)) : source.name;
			const normalizedName = name.toLocaleLowerCase("en-US");
			if (names.has(normalizedName)) {
				throw new Error(`Duplicate Power Query source name "${name}".`);
			}
			names.add(normalizedName);

			const formula = (await readFile(file, "utf8")).replace(/^\uFEFF/, "");
			queries.push({ name, formula });
		}
	}

	return queries;
}

function resolvePowerQueryFiles(project: Project, source: Source): string[] {
	if (!source.path.includes("*")) return [source.path];

	const baseDir = project.paths.dir;
	const normalizedSourcePath = source.path.replace(/\\/g, "/");
	const normalizedBaseDir = baseDir.replace(/\\/g, "/");
	const pattern = normalizedSourcePath.startsWith(normalizedBaseDir)
		? relative(baseDir, source.path)
		: source.path;

	return walk(baseDir, { globs: [pattern], directories: false })
		.filter(isPowerQueryPath)
		.map(file => join(baseDir, file));
}
