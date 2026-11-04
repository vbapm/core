import { copyFile, mkdir, readdir, rm } from "fs/promises";
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

export async function exportPowerQueries(project: Project, staging: string): Promise<void> {
	const stagedDirectory = nativeJoin(staging, "pq");
	let stagedFiles;
	try {
		stagedFiles = await readdir(stagedDirectory, { withFileTypes: true });
	} catch (error) {
		if ((error as NodeJS.ErrnoException).code === "ENOENT") return;
		throw error;
	}

	const subfolders = project.manifest.srcProperties?.subfolders;
	const hasSubfolders =
		subfolders != null && Object.values(subfolders).some(folder => folder.trim().length > 0);
	const powerQuerySources = project.manifest.src.filter(isPowerQuerySource);
	const individualSources = powerQuerySources.filter(source => !source.path.includes("*"));
	const wildcardSources = powerQuerySources.filter(source => source.path.includes("*"));
	const queryFolder = hasSubfolders ? subfolders.PowerQueries?.trim() || "Power Queries" : "pq";
	const defaultDirectory = nativeJoin(
		project.paths.dir,
		resolveSrcFolder(project.manifest.srcProperties),
		queryFolder
	);

	const queries = stagedFiles.filter(
		file => file.isFile() && extname(file.name).toLowerCase() === ".pq"
	);
	const queryNames = new Set(
		queries.map(file => sourceName(basename(file.name, extname(file.name))))
	);
	const wildcardMatches = wildcardSources.map(source => ({
		source,
		files: expandPowerQuerySource(project.paths.dir, source.path)
	}));
	const explicitlyMappedPaths = new Set(
		individualSources.map(source => resolveSourcePath(project.paths.dir, source.path))
	);

	for (const source of individualSources) {
		if (!queryNames.has(sourceName(source.name))) {
			await rm(resolveSourcePath(project.paths.dir, source.path), { force: true });
		}
	}

	for (const mapping of wildcardMatches) {
		for (const file of mapping.files) {
			if (
				!queryNames.has(sourceName(basename(file, extname(file)))) &&
				!explicitlyMappedPaths.has(resolve(file))
			) {
				await rm(file);
			}
		}
	}
	if (wildcardMatches.length === 0 && individualSources.length === 0) {
		try {
			const existingFiles = await readdir(defaultDirectory, { withFileTypes: true });
			for (const file of existingFiles) {
				if (
					file.isFile() &&
					extname(file.name).toLowerCase() === ".pq" &&
					!queryNames.has(sourceName(basename(file.name, extname(file.name))))
				) {
					await rm(nativeJoin(defaultDirectory, file.name));
				}
			}
		} catch (error) {
			if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
		}
	}

	for (const file of queries) {
		const name = sourceName(basename(file.name, extname(file.name)));
		const individualSource = individualSources.find(source => sourceName(source.name) === name);
		const existingWildcardMatch = wildcardMatches
			.flatMap(mapping => mapping.files)
			.find(existingFile => sourceName(basename(existingFile, extname(existingFile))) === name);
		const wildcardRoot =
			wildcardMatches[0] &&
			resolveSourcePath(project.paths.dir, resolveWildcardRoot(wildcardMatches[0].source.path));
		const destination =
			(individualSource && resolveSourcePath(project.paths.dir, individualSource.path)) ??
			existingWildcardMatch ??
			nativeJoin(wildcardRoot || defaultDirectory, file.name);
		await mkdir(dirname(destination), { recursive: true });
		await copyFile(nativeJoin(stagedDirectory, file.name), destination);
	}

	const hasUnmappedQueries = queries.some(query => {
		const name = sourceName(basename(query.name, extname(query.name)));
		return !individualSources.some(source => sourceName(source.name) === name);
	});
	if (hasUnmappedQueries && hasSubfolders && wildcardSources.length === 0) {
		project.manifest.srcProperties!.subfolders = { ...subfolders, PowerQueries: queryFolder };

		const querySourcePath = nativeJoin(
			project.paths.dir,
			resolveSrcFolder(project.manifest.srcProperties),
			queryFolder,
			"*.pq"
		);
		const sourceName = availableSourceName(project.manifest.src, "PowerQueries");
		project.manifest.src.push({ name: sourceName, path: querySourcePath });

		await writeManifest(project.manifest, project.paths.dir);
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

function resolveWildcardRoot(sourcePath: string): string {
	const wildcardIndex = sourcePath.search(/[?*]/);
	if (wildcardIndex < 0) return dirname(sourcePath);

	const prefix = sourcePath.slice(0, wildcardIndex);
	const lastSeparator = Math.max(prefix.lastIndexOf("/"), prefix.lastIndexOf("\\"));
	return prefix.endsWith("/") || prefix.endsWith("\\")
		? prefix.slice(0, -1)
		: prefix.slice(0, lastSeparator);
}
