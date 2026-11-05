import { copyFile, mkdir, readFile, readdir, rm } from "fs/promises";
import walk from "walk-sync";
import {
	basename,
	dirname,
	extname,
	isAbsolute,
	join as nativeJoin,
	relative,
	resolve
} from "path";
import { resolveSrcFolder, writeManifest } from "../manifest";
import { isPowerQuerySource } from "../manifest/source";
import { Project } from "../project";

interface StagedQuery {
	name: string;
	file: string;
}

interface QueryMapping {
	source: Project["manifest"]["src"][number];
	files: string[];
}

export async function exportPowerQueries(project: Project, staging: string): Promise<void> {
	const stagedDirectory = nativeJoin(staging, "pq");
	let stagedFiles;
	try {
		stagedFiles = await readdir(stagedDirectory, { withFileTypes: true });
	} catch (error) {
		if ((error as NodeJS.ErrnoException).code === "ENOENT") return;
		throw error;
	}

	const stagedNames = await readStagedQueryNames(staging);
	const queries: StagedQuery[] = stagedFiles
		.filter(file => file.isFile() && extname(file.name).toLowerCase() === ".pq")
		.map(file => ({
			file: file.name,
			name: stagedNames.get(sourceName(file.name)) ?? basename(file.name, extname(file.name))
		}));
	const queryNames = new Set(queries.map(query => sourceName(query.name)));
	const srcSources = project.manifest.src.filter(isPowerQuerySource);
	const devSources = (project.manifest.devSrc ?? []).filter(isPowerQuerySource);
	const srcIndividuals = srcSources.filter(source => !source.path.includes("*"));
	const srcWildcards = getWildcardMappings(project.paths.dir, srcSources);
	const devIndividuals = devSources.filter(source => !source.path.includes("*"));
	const devWildcards = getWildcardMappings(project.paths.dir, devSources);
	const subfolders = project.manifest.srcProperties?.subfolders;
	const hasSubfolders =
		subfolders != null && Object.values(subfolders).some(folder => folder.trim().length > 0);
	const queryFolder = hasSubfolders ? subfolders.PowerQueries?.trim() || "Power Queries" : "pq";
	const defaultDirectory = nativeJoin(
		project.paths.dir,
		resolveSrcFolder(project.manifest.srcProperties),
		queryFolder
	);
	const routes = queries.map(query =>
		resolveQueryRoute(
			project.paths.dir,
			query,
			devIndividuals,
			devWildcards,
			srcIndividuals,
			srcWildcards,
			defaultDirectory
		)
	);
	const retainedPaths = new Set(routes.map(route => pathKey(route.destination)));
	const devOwnedPaths = new Set([
		...devIndividuals.map(source => pathKey(resolveSourcePath(project.paths.dir, source.path))),
		...devWildcards.flatMap(mapping => mapping.files.map(pathKey))
	]);
	let manifestChanged = false;

	const retainedSources = [];
	for (const source of project.manifest.src) {
		if (!isPowerQuerySource(source) || source.path.includes("*")) {
			retainedSources.push(source);
			continue;
		}
		const sourcePath = resolveSourcePath(project.paths.dir, source.path);
		if (queryNames.has(sourceName(source.name)) || retainedPaths.has(pathKey(sourcePath))) {
			retainedSources.push(source);
			continue;
		}
		if (!devOwnedPaths.has(pathKey(sourcePath))) await rm(sourcePath, { force: true });
		manifestChanged = true;
	}
	if (retainedSources.length !== project.manifest.src.length) {
		project.manifest.src = retainedSources;
	}

	for (const mapping of srcWildcards) {
		for (const file of mapping.files) {
			const key = pathKey(file);
			if (!retainedPaths.has(key) && !devOwnedPaths.has(key)) await rm(file, { force: true });
		}
	}
	if (srcWildcards.length === 0 && srcIndividuals.length === 0) {
		try {
			const files = await readdir(defaultDirectory, { withFileTypes: true });
			for (const file of files) {
				const path = nativeJoin(defaultDirectory, file.name);
				if (
					file.isFile() &&
					extname(file.name).toLowerCase() === ".pq" &&
					!retainedPaths.has(pathKey(path)) &&
					!devOwnedPaths.has(pathKey(path))
				) {
					await rm(path, { force: true });
				}
			}
		} catch (error) {
			if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
		}
	}

	let groupRegistered = false;
	for (const route of routes) {
		await mkdir(dirname(route.destination), { recursive: true });
		await copyFile(nativeJoin(stagedDirectory, route.query.file), route.destination);
		const filenameMatchesName =
			sourceName(basename(route.query.file, extname(route.query.file))) ===
			sourceName(route.query.name);
		if (route.explicit || (route.mapped && filenameMatchesName)) continue;

		const sourceSection = route.development ? project.manifest.devSrc : project.manifest.src;
		if (!filenameMatchesName) {
			sourceSection.push({ name: route.query.name, path: route.destination });
			manifestChanged = true;
			continue;
		}

		const coveredByWildcard = srcWildcards.some(mapping =>
			expandPowerQuerySource(project.paths.dir, mapping.source.path).some(
				file => pathKey(file) === pathKey(route.destination)
			)
		);
		if (coveredByWildcard) continue;

		if (srcWildcards.length === 0 && hasSubfolders) {
			if (groupRegistered) continue;
			project.manifest.srcProperties!.subfolders = { ...subfolders, PowerQueries: queryFolder };
			project.manifest.src.push({
				name: availableSourceName(project.manifest.src, "PowerQueries"),
				path: nativeJoin(
					project.paths.dir,
					resolveSrcFolder(project.manifest.srcProperties),
					queryFolder,
					"*.pq"
				)
			});
			groupRegistered = true;
			manifestChanged = true;
			continue;
		}

		project.manifest.src.push({ name: route.query.name, path: route.destination });
		manifestChanged = true;
	}

	if (manifestChanged) await writeManifest(project.manifest, project.paths.dir);
}

interface QueryRoute {
	query: StagedQuery;
	destination: string;
	mapped: boolean;
	explicit: boolean;
	development: boolean;
}

function resolveQueryRoute(
	baseDir: string,
	query: StagedQuery,
	devIndividuals: Project["manifest"]["src"],
	devWildcards: QueryMapping[],
	srcIndividuals: Project["manifest"]["src"],
	srcWildcards: QueryMapping[],
	defaultDirectory: string
): QueryRoute {
	const name = sourceName(query.name);
	const individual =
		devIndividuals.find(source => sourceName(source.name) === name) ??
		srcIndividuals.find(source => sourceName(source.name) === name);
	if (individual) {
		return {
			query,
			destination: resolveSourcePath(baseDir, individual.path),
			mapped: true,
			explicit: true,
			development: devIndividuals.includes(individual)
		};
	}

	const matchingWildcard = [...devWildcards, ...srcWildcards].find(mapping =>
		mapping.files.some(file => isQueryFileMatch(file, query))
	);
	if (matchingWildcard) {
		const destination = matchingWildcard.files.find(file => isQueryFileMatch(file, query))!;
		return {
			query,
			destination,
			mapped: true,
			explicit: false,
			development: devWildcards.includes(matchingWildcard)
		};
	}

	const sourceWildcard = srcWildcards[0];
	const directory = sourceWildcard
		? resolveSourcePath(baseDir, resolveWildcardRoot(sourceWildcard.source.path))
		: defaultDirectory;
	return {
		query,
		destination: nativeJoin(directory, query.file),
		mapped: false,
		explicit: false,
		development: false
	};
}

function isQueryFileMatch(file: string, query: StagedQuery): boolean {
	const fileName = sourceName(basename(file, extname(file)));
	return (
		fileName === sourceName(basename(query.file, extname(query.file))) ||
		fileName === sourceName(query.name)
	);
}

function getWildcardMappings(baseDir: string, sources: Project["manifest"]["src"]): QueryMapping[] {
	return sources
		.filter(source => source.path.includes("*"))
		.map(source => ({ source, files: expandPowerQuerySource(baseDir, source.path) }));
}

async function readStagedQueryNames(staging: string): Promise<Map<string, string>> {
	try {
		const projectInfo = JSON.parse(await readFile(nativeJoin(staging, "project.json"), "utf8"));
		return new Map(
			(projectInfo.powerQueries ?? [])
				.filter((query: any) => typeof query.name === "string" && typeof query.file === "string")
				.map((query: { name: string; file: string }) => [sourceName(query.file), query.name])
		);
	} catch (error) {
		if ((error as NodeJS.ErrnoException).code === "ENOENT") return new Map();
		throw error;
	}
}

function sourceName(value: string): string {
	return value.toLocaleLowerCase("en-US");
}

function availableSourceName(sources: Project["manifest"]["src"], baseName: string): string {
	const existingNames = new Set(sources.map(source => sourceName(source.name)));
	let candidate = baseName;
	let suffix = 2;
	while (existingNames.has(sourceName(candidate))) {
		candidate = `${baseName}${suffix++}`;
	}
	return candidate;
}

function expandPowerQuerySource(baseDir: string, sourcePath: string): string[] {
	const normalizedSourcePath = sourcePath.replace(/\\/g, "/");
	const normalizedBaseDir = baseDir.replace(/\\/g, "/");
	const pattern = normalizedSourcePath.startsWith(normalizedBaseDir)
		? relative(baseDir, sourcePath)
		: sourcePath;
	return walk(baseDir, { globs: [pattern], directories: false })
		.filter(file => extname(file).toLowerCase() === ".pq")
		.map(file => nativeJoin(baseDir, file));
}

function resolveSourcePath(baseDir: string, sourcePath: string): string {
	return isAbsolute(sourcePath) ? sourcePath : resolve(baseDir, sourcePath);
}

function pathKey(path: string): string {
	const resolved = resolve(path);
	return process.platform === "win32" ? resolved.toLocaleLowerCase("en-US") : resolved;
}

function resolveWildcardRoot(sourcePath: string): string {
	const wildcardIndex = sourcePath.search(/[?*]/);
	if (wildcardIndex < 0) return dirname(sourcePath);

	const prefix = sourcePath.slice(0, wildcardIndex);
	const lastSeparator = Math.max(prefix.lastIndexOf("/"), prefix.lastIndexOf("\\"));
	return prefix.endsWith("/") || prefix.endsWith("\\")
		? prefix.slice(0, -1)
		: prefix.slice(0, lastSeparator);
}
