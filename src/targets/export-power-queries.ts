import { copyFile, mkdir, readdir, rm } from "fs/promises";
import { dirname, extname, join as nativeJoin } from "path";
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
	const existingPowerQuerySource = project.manifest.src.find(isPowerQuerySource);
	const queryFolder = hasSubfolders ? subfolders.PowerQueries?.trim() || "Power Queries" : "pq";
	const sourceDirectory = existingPowerQuerySource
		? dirname(existingPowerQuerySource.path)
		: nativeJoin(project.paths.dir, resolveSrcFolder(project.manifest.srcProperties), queryFolder);
	await mkdir(sourceDirectory, { recursive: true });

	const queries = stagedFiles.filter(
		file => file.isFile() && extname(file.name).toLowerCase() === ".pq"
	);
	const queryNames = new Set(queries.map(file => file.name));
	const existingFiles = await readdir(sourceDirectory, { withFileTypes: true });

	for (const file of existingFiles) {
		if (file.isFile() && extname(file.name).toLowerCase() === ".pq" && !queryNames.has(file.name)) {
			await rm(nativeJoin(sourceDirectory, file.name));
		}
	}

	for (const file of queries) {
		await copyFile(nativeJoin(stagedDirectory, file.name), nativeJoin(sourceDirectory, file.name));
	}

	if (queries.length > 0 && hasSubfolders && !existingPowerQuerySource) {
		project.manifest.srcProperties!.subfolders = { ...subfolders, PowerQueries: queryFolder };

		const querySourcePath = nativeJoin(
			project.paths.dir,
			resolveSrcFolder(project.manifest.srcProperties),
			queryFolder,
			"*.pq"
		);
		const existingSource = project.manifest.src.find(source => source.name === "PowerQueries");
		if (existingSource) {
			existingSource.path = querySourcePath;
		} else {
			project.manifest.src.push({ name: "PowerQueries", path: querySourcePath });
		}

		await writeManifest(project.manifest, project.paths.dir);
	}
}
