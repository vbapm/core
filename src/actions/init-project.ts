import dedent from "@timhall/dedent";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { Manifest, writeManifest } from "../manifest";
import { TargetType } from "../manifest/target";
import { initProject as init } from "../project";
import { addTarget } from "../targets/add-target";
import { copy, ensureDir, pathExists } from "../utils/fs";
import { init as git_init } from "../utils/git";
import { basename, extname, join } from "../utils/path";

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
}

export async function initProject(options: InitOptions) {
	let { name, dir = env.cwd, target: targetType, from, pkg: asPackage, git, configTemplates } = options;

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

	if (from) {
		targetType = extname(from).replace(".", "");
	}
	if (targetType) {
		const dependencies: Manifest[] = [];
		await addTarget(<TargetType>targetType, { project, dependencies }, { from });
	}

	await writeManifest(project.manifest, project.paths.dir);
}
