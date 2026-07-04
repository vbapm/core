import dedent from "@timhall/dedent";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { Manifest, writeManifest } from "../manifest";
import { TargetType } from "../manifest/target";
import { initProject as init } from "../project";
import { addTarget } from "../targets/add-target";
import { copy, ensureDir, pathExists, readFile } from "../utils/fs";
import { init as git_init } from "../utils/git";
import { basename, extname, join } from "../utils/path";
import { codepageToLabel, getSystemCodepage } from "../build/encoding-sniffer";

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
	let {
		name,
		dir = env.cwd,
		target: targetType,
		from,
		pkg: asPackage,
		git,
		configTemplates
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

	await writeManifest(project.manifest, project.paths.dir);
}

/**
 * Auto-detect the source encoding for a project initialized from an
 * existing workbook. If any source file contains non-ASCII characters,
 * the system codepage is used as the default. If jschardet confidently
 * identifies a different encoding (e.g. a Japanese workbook opened on
 * a Western European system), that encoding is used instead.
 */
async function detectSourceEncoding(project: import("../project").Project): Promise<void> {
	// Check if any source file has non-ASCII characters
	let hasNonAscii = false;
	for (const source of project.manifest.src) {
		try {
			const code = (await readFile(join(project.paths.dir, source.path))).toString();
			if (/[^\x00-\x7F]/.test(code)) {
				hasNonAscii = true;
				break;
			}
		} catch {
			// File may not be readable — skip
		}
	}

	if (!hasNonAscii) return;

	// Determine encoding: default to system codepage
	const systemCp = getSystemCodepage();
	const systemLabel = codepageToLabel(systemCp);
	let srcEncoding = systemLabel;

	// Try jschardet to see if a different encoding is more likely
	try {
		const jschardet = require("jschardet");
		const SUPPORTED = /^(CP(932|936|949|950|874|125[0-8]))$/;

		// Read the first source file as a sample for detection
		const firstSource = project.manifest.src[0];
		if (firstSource) {
			const buffer = await readFile(join(project.paths.dir, firstSource.path));
			const results = (jschardet.detectAll(buffer) as Array<{ encoding: string; confidence: number }>)
				.filter((r: { encoding: string }) => SUPPORTED.test(r.encoding))
				.sort((a: { confidence: number }, b: { confidence: number }) => b.confidence - a.confidence);

			// Use jschardet's guess only if it's confident AND differs from system codepage
			const SYSTEM_PREFIX = /^windows-?/i;
			if (results.length > 0 && results[0].confidence >= 0.8) {
				const detected = results[0].encoding.toLowerCase();
				// jschardet returns "CP1252" style; convert to "cp1252"
				const detectedLabel = detected.toLowerCase();
				const systemNormalized = systemLabel.replace(SYSTEM_PREFIX, "cp").toLowerCase();

				if (detectedLabel !== systemNormalized) {
					srcEncoding = detectedLabel;
				}
			}
		}
	} catch {
		// jschardet unavailable — keep system codepage
	}

	project.manifest.srcEncoding = srcEncoding;

	// Set target encoding to match the detected encoding, so the
	// workbook's VBA encoding is preserved when building.
	if (project.manifest.target) {
		project.manifest.target.encoding = srcEncoding;
	}
}
