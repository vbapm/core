import dedent from "@timhall/dedent";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { Manifest, writeManifest } from "../manifest";
import { TargetType } from "../manifest/target";
import { initProject as init } from "../project";
import { addTarget } from "../targets/add-target";
import { copy, ensureDir, pathExists } from "../utils/fs";
import { init as git_init } from "../utils/git";
import { basename, dirname, extname, join } from "../utils/path";
import { detectImportEncoding } from "./detect-encoding";

const TEMPLATE_FILES = [
	{ source: "template.gitignore", target: ".gitignore" },
	{ source: "template.gitattributes", target: ".gitattributes" },
	{ source: "template.editorconfig", target: ".editorconfig" }
];

async function copyTemplateConfigFiles(dir: string) {
	const templatesDir = join(__dirname, "templates");

	for (const { source, target } of TEMPLATE_FILES) {
		await copy(join(templatesDir, source), join(dir, target));
	}
}

export interface InitOptions {
	name?: string;
	dir?: string;
	target?: string;
	from?: string;
	pkg: boolean;
	git: boolean;
	configTemplates: boolean;
	individual?: boolean;
}

export async function initProject(options: InitOptions) {
	let {
		name,
		dir = env.cwd,
		target: targetType,
		from,
		pkg: asPackage,
		git,
		configTemplates,
		listAll
	} = options;

	if (await pathExists(join(dir, "vbaproject.toml"))) {
		throw new CliError(
			ErrorCode.InitAlreadyInitialized,
			`A vbapm project already exists in this directory.`
		);
	}

	if (from && !(await pathExists(from))) {
		throw new CliError(ErrorCode.FromNotFound, `The "from" document was not found at "${from}".`);
	}

	name = name || (from ? basename(from, extname(from)) : basename(dir));

	if (!name) {
		throw new CliError(
			ErrorCode.InitNameRequired,
			dedent`
        Unable to determine name from current directory or --from.
        --name NAME is required to initialize this project.
      `
		);
	}
	if (!targetType && !from && name.includes(".")) {
		const parts = name.split(".");
		targetType = parts.pop();
		name = parts.join(".");
	}

	if (!asPackage && !targetType && !from) {
		throw new CliError(
			ErrorCode.InitTargetRequired,
			dedent`
        --target or --from is required for vbapm projects.
        (e.g. vbapm init --target xlsm)
      `
		);
	}

	await ensureDir(join(dir, "src"));

	if (configTemplates) {
		await copyTemplateConfigFiles(dir);
	}

	if (git && !(await pathExists(join(dir, ".git")))) {
		await git_init(dir);
	}

	const project = await init(name, dir, {
		type: asPackage ? "package" : "project"
	});

	// When importing from a workbook at the project root, default build-dir to "."
	// so the built file is written alongside the source workbook
	if (from && dirname(from) === dir) {
		project.manifest.buildDir = ".";
		project.paths.build = join(dir, ".");
		project.paths.backup = join(dir, ".", ".backup");
	}

	if (from) {
		targetType = extname(from).replace(".", "");
	}
	if (targetType) {
		const dependencies: Manifest[] = [];
		await addTarget(<TargetType>targetType, { project, dependencies }, { from });
	}

	// Auto-detect src-encoding when importing from an existing workbook
	if (from) {
		await detectSourceEncoding(project);
	}

	// Write default wildcard [src] entries for new projects
	if (!listAll) {
		project.manifest.src = [
			{ name: "Modules", path: "src/**/*.bas" },
			{ name: "Forms", path: "src/**/*.frm" },
			{ name: "Classes", path: "src/**/*.cls" }
		];
	}

	await writeManifest(project.manifest, project.paths.dir);
}

/**
 * Auto-detect the source encoding for a project initialized from an
 * existing workbook.
 *
 * Important: this should be used right after the first export of the workbook to src,
 * before any source files are modified.
 *
 * If any source file contains non-ASCII characters,
 * the system codepage is used as the default. If jschardet confidently
 * identifies a different encoding (e.g. a Japanese workbook opened on
 * a Western European system), that encoding is used instead.
 */
async function detectSourceEncoding(project: import("../project").Project): Promise<void> {
	const firstSource = project.manifest.src[0];
	if (!firstSource) return;

	const encoding = await detectImportEncoding(join(project.paths.dir, firstSource.path));
	if (!encoding) return;

	project.manifest.srcEncoding = encoding;
	if (project.manifest.target) {
		project.manifest.target.encoding = encoding;
	}
}
