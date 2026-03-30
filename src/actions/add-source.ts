import dedent from "@timhall/dedent";
import { env } from "../env";
import { CliError, ErrorCode } from "../errors";
import { loadManifest, writeManifest } from "../manifest";
import { Source } from "../manifest/source";
import { ensureDir, pathExists, writeFile } from "../utils/fs";
import { basename, extname, join } from "../utils/path";

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

export async function addSource(options: AddSourceOptions): Promise<string> {
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
	if (name.includes("/") || name.includes("\\")) {
		throw new CliError(
			ErrorCode.AddInvalidName,
			`Invalid source name "${name}". Use a VBA component name without path separators.`
		);
	}

	const extension = getExtension(name, type);
	const componentName = basename(name, extname(name));
	if (!componentName) {
		throw new CliError(ErrorCode.AddInvalidName, `Invalid source name "${name}".`);
	}

	const filename = `${componentName}${extension}`;
	const path = join(root, "src", filename);

	const section = dev ? manifest.devSrc : manifest.src;
	const duplicate = section.find(source => source.name === componentName || source.path === path);
	if (duplicate) {
		throw new CliError(
			ErrorCode.AddSourceExists,
			`Source "${componentName}" already exists in [${dev ? "dev-src" : "src"}].`
		);
	}

	if (await pathExists(path)) {
		throw new CliError(
			ErrorCode.AddSourceExists,
			`A file already exists at "${path}". Remove it first or choose a different name.`
		);
	}

	await ensureDir(join(root, "src"));
	await writeFile(path, template(componentName, extension));

	const source: Source = {
		name: componentName,
		path
	};
	section.push(source);

	await writeManifest(manifest, root);

	return path;
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
		return dedent`
      VERSION 1.0 CLASS
      BEGIN
        MultiUse = -1  'True
      END
      Attribute VB_Name = "${name}"
      Attribute VB_GlobalNameSpace = False
      Attribute VB_Creatable = False
      Attribute VB_PredeclaredId = False
      Attribute VB_Exposed = False

    `;
	}

	return dedent`
    Attribute VB_Name = "${name}"

  `;
}
