import dedent from "@timhall/dedent";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { loadManifest, writeManifest } from "../manifest";
import { Source } from "../manifest/source";
import { ensureDir, pathExists, writeFile } from "../utils/fs";
import { basename, dirname, extname, join } from "../utils/path";

const extensionToType: { [extension: string]: "module" | "class" } = {
	".bas": "module",
	".cls": "class"
};

const typeToExtension: { [type: string]: ".bas" | ".cls" } = {
	module: ".bas",
	class: ".cls"
};

export interface AddSourceOptions {
	name?: string;
	type?: string;
	dir?: string;
	dev: boolean;
}

export interface AddSourceResult {
	path: string;
	isNew: boolean;
}

export async function addSource(options: AddSourceOptions): Promise<AddSourceResult> {
	const { name, type, dir, dev } = options;
	const root = dir || env.cwd;

	if (!name) {
		throw new CliError(
			ErrorCode.AddNameRequired,
			dedent`
        "name" is required with vbapm add (e.g. vbapm add Module1).

        Try "vbapm help add" for more information.
      `
		);
	}

	const manifest = await loadManifest(root);
	const extension = getExtension(name, type);
	const hasPathParts = name.includes("/") || name.includes("\\");
	let path = hasPathParts ? join(root, name) : join(root, "src", name);
	if (!extname(path)) {
		path = `${path}${extension}`;
	}

	const componentName = basename(path, extname(path));
	if (!componentName) {
		throw new CliError(ErrorCode.AddInvalidName, `Invalid source name "${name}".`);
	}

	const sectionName = dev ? "dev-src" : "src";
	const existingInSrc = manifest.src.find(
		source => source.name === componentName || source.path === path
	);
	const existingInDevSrc = manifest.devSrc.find(
		source => source.name === componentName || source.path === path
	);
	if (existingInSrc || existingInDevSrc) {
		const existingSection = existingInSrc ? "src" : "dev-src";
		throw new CliError(
			ErrorCode.AddSourceExists,
			`Source "${componentName}" already exists in [${existingSection}] and cannot be added to [${sectionName}].`
		);
	}

	const section = dev ? manifest.devSrc : manifest.src;

	const fileExists = await pathExists(path);

	if (!fileExists) {
		await ensureDir(dirname(path));
		await writeFile(path, template(componentName, extension));
	}

	const source: Source = {
		name: componentName,
		path
	};
	section.push(source);

	await writeManifest(manifest, root);

	return { path, isNew: !fileExists };
}

function getExtension(name: string, type?: string): ".bas" | ".cls" {
	const fromName = extname(name).toLowerCase();
	const fromType = type && typeToExtension[type.toLowerCase()];

	if (fromName && !(fromName in extensionToType)) {
		throw new CliError(
			ErrorCode.AddUnsupportedType,
			`Unsupported source extension "${fromName}". Supported extensions are .bas and .cls.`
		);
	}

	if (type && !fromType) {
		throw new CliError(
			ErrorCode.AddUnsupportedType,
			`Unsupported --type "${type}". Supported types are "module" and "class".`
		);
	}

	if (fromName && fromType && fromName !== fromType) {
		throw new CliError(
			ErrorCode.AddUnsupportedType,
			`Type "${type}" conflicts with extension "${fromName}".`
		);
	}

	if (fromName) return <".bas" | ".cls">fromName;
	if (fromType) return fromType;

	return ".bas";
}

function template(name: string, extension: ".bas" | ".cls"): string {
	if (extension === ".cls") {
		return (
			dedent`
      VERSION 1.0 CLASS
      BEGIN
        MultiUse = -1  'True
      END
      Attribute VB_Name = "${name}"
      Attribute VB_GlobalNameSpace = False
      Attribute VB_Creatable = False
      Attribute VB_PredeclaredId = False
      Attribute VB_Exposed = False
    ` + "\n\n"
		);
	}

	return (
		dedent`
    Attribute VB_Name = "${name}"
    ` + "\n\n"
	);
}
