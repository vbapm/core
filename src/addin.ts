import { ImportGraph } from "./build/build-graph";
import { readdir } from "fs/promises";
import { env } from "./env";
import { CliError, ErrorCode } from "./errors";
import { resolvePeerReferencePaths } from "./manifest/reference";
import { Target } from "./manifest/target";
import { Project } from "./project";
import { copy, ensureDir, pathExists } from "./utils/fs";
import { dirname, extname, join } from "./utils/path";
import { run } from "./utils/run";

export type Application = "excel";
export type Addin = string;

export interface AddinOptions {
	addin?: string;
	open?: boolean;
	staging?: boolean;
}

export const extensions: { [application: string]: string[] } = {
	excel: ["xlsx", "xlsm", "xlam"]
};
export const addins: { [application: string]: string } = {
	excel: join(env.addins, "vbapm.xlam")
};

const byExtension: { [extension: string]: string } = {};
for (const [application, values] of Object.entries(extensions)) {
	for (const extension of values) {
		byExtension[extension] = application;
	}
}

/**
 * Import graph of src and references into given target
 */
export async function importGraph(
	project: Project,
	target: Target,
	graph: ImportGraph,
	file: string,
	options: AddinOptions = {}
): Promise<void> {
	const { application, addin } = getTargetInfo(project, target);
	const { name, components, references } = graph;

	// Resolve peer (VBA project) reference paths to absolute before sending to
	// the addin. The manifest stores them relative (when nearby) or absolute;
	// `References.AddFromFile` needs an absolute path.
	const resolvedReferences = resolvePeerReferencePaths(references, project.paths.dir);
	for (const reference of references) {
		if (reference.peer && !reference.path) {
			throw new CliError(
				ErrorCode.PeerReferenceMissingPath,
				`Peer reference <${reference.name}> is missing a "path".\n\n` +
					`Add the path to the addin in vbaproject.toml, e.g.:\n\n` +
					`  [references.${reference.name}]\n` +
					`  peer = true\n` +
					`  path = "../${reference.name}/${reference.name}.xlam"`
			);
		}
	}

	await run(
		application,
		options.addin || addin,
		"Build.ImportGraph",
		[
			JSON.stringify({
				file,
				name,
				src: components,
				references: resolvedReferences
			})
		],
		{ keepOpen: options.open }
	);
}

/**
 * Export src and references from given target
 */
export async function exportTo(
	project: Project,
	target: Target,
	staging: string,
	options: AddinOptions = {}
): Promise<void> {
	let { application, addin, file } = getTargetInfo(project, target);
	await validateExportTarget(file, target, project.manifest.name);

	// For Mac, stage target to avoid permission prompts
	if (!env.isWindows) {
		const staged = join(staging, "staged", target.filename);
		if (!(await pathExists(staged))) {
			await ensureDir(dirname(staged));
			await copy(file, staged);
		}

		file = staged;
	}

	await run(application, options.addin || addin, "Build.ExportTo", [
		JSON.stringify({
			file,
			staging,
			// Respect [source] "include-empty-objects" flag (default: true)
			includeEmptyObjects: project.manifest.srcProperties?.["include-empty-objects"] ?? true
		})
	]);
}

export async function validateExportTarget(
	file: string,
	target: Target,
	projectName: string
): Promise<void> {
	if (await pathExists(file)) return;

	const directory = dirname(file);
	if (!(await pathExists(directory))) {
		let missingDirectory = directory;
		let parent = dirname(missingDirectory);

		while (parent !== missingDirectory && !(await pathExists(parent))) {
			missingDirectory = parent;
			parent = dirname(missingDirectory);
		}

		throw new CliError(
			ErrorCode.ExportTargetNotFound,
			`The directory containing the target file does not exist:\n\n  "${missingDirectory}"`
		);
	}

	const extension = extname(file).toLowerCase();
	const candidates = (await readdir(directory, { withFileTypes: true }))
		.filter(entry => entry.isFile() && extname(entry.name).toLowerCase() === extension)
		.map(entry => entry.name)
		.sort((left, right) => left.localeCompare(right));

	let message = `The target file does not exist:\n\n  "${file}"`;
	if (candidates.length > 0) {
		message += `\n\nFound other ${extension} files in "${directory}":\n${candidates
			.map(candidate => `  - "${candidate}"`)
			.join("\n")}`;

		const suggestedName =
			candidates.length === 1 ? candidates[0].slice(0, -extension.length) : "WORKBOOK_NAME";
		message +=
			`\n\nThe target filename defaults to the project name "${projectName}". ` +
			`If the workbook name differs, set target.name in vbaproject.toml without the extension:\n\n` +
			`  [project]\n  target = { type = "${target.type}", name = "${suggestedName}" }`;
	}

	throw new CliError(ErrorCode.ExportTargetNotFound, message);
}

/**
 * Create a new document at the given path
 */
export async function createDocument(
	project: Project,
	target: Target,
	options: AddinOptions = {}
): Promise<string> {
	const { application, addin, file } = getTargetInfo(project, target, options);

	// For Mac, stage target to avoid permission prompts and then copy to build directory
	const useStaging = !env.isWindows && !options.staging;
	let path = !useStaging ? file : join(project.paths.staging, target.filename);

	await ensureDir(dirname(path));
	await run(application, options.addin || addin, "Build.CreateDocument", [
		JSON.stringify({
			path
		})
	]);

	// For Mac, then copy staged to build directory
	if (useStaging) {
		await ensureDir(dirname(file));
		await copy(path, file);
	}

	return file;
}

/**
 * Get application, addin, and file for given target
 */
export function getTargetInfo(
	project: Project,
	target: Target,
	options: AddinOptions = {}
): { application: Application; addin: Addin; file: string } {
	const application = extensionToApplication(target.type);
	const addin = addins[application];
	const file = join(options.staging ? project.paths.staging : project.paths.build, target.filename);

	return { application, addin, file };
}

export function extensionToApplication(extension: string): Application {
	extension = extension.replace(".", "");
	const application = byExtension[extension];
	if (!application) {
		throw new CliError(
			ErrorCode.AddinUnsupportedType,
			`The target type "${extension} is not currently supported.`
		);
	}

	return application as Application;
}
