import dedent from "@timhall/dedent";
import { CliError, ErrorCode, manifestOk } from "../errors";
import { pathExists, readFile, writeFile } from "../utils/fs";
import { join, normalize } from "../utils/path";
import { convert as convertToToml, parse as parseToml, patch as patchToml } from "../utils/toml";
import { Dependency, formatDependencies, parseDependencies } from "./dependency";
import { formatReferences, parseReferences, Reference } from "./reference";
import { formatSrc, parseSrc, Source } from "./source";
import {
	detectSrcStructure,
	parseSrcProperties,
	resolveSrcFolder,
	resolveSrcSubfolders,
	SrcProperties,
	SrcStructure,
	SrcSubfolders
} from "./src-sort";
import { formatTarget, parseTarget, Target } from "./target";
import { DEFAULT_VERSION, Version } from "./version";

// Re-export for consumers
export { resolveSrcFolder, resolveSrcSubfolders, SrcSubfolders };

/**
 * Snapshot is the minimal manifest needed to support both Manifest
 * and info loaded during dependency resolution
 */
export interface Snapshot {
	name: string;
	version: Version;
	dependencies: Dependency[];
}

export type ManifestType = "package" | "project";

/*
  # Manifest

  The parsed vbaproject.toml manifest.
  package/project, src, dependencies, etc are all parsed and put in a consistent form

  ```toml
  [package]
  name = "package-name"
  version = "1.0.0"
  authors = ["Tim Hall <tim.hall.engr@gmail.com> (https://github.com/timhall)"]

  [src]
  A = "src/a.bas"
  B = { path = "src/b.cls" }

  [dependencies]
  dictionary = "^1"
  from-path = { path = "packages/from-path" }
```
*/
export interface Metadata {
	authors?: string[];
	publish?: boolean;
	[name: string]: any;
}

export interface Manifest extends Snapshot {
	type: ManifestType;
	metadata: Metadata;
	src: Source[];
	srcEncoding?: string;
	srcProperties?: SrcProperties;
	srcStructure?: SrcStructure;
	codename?: string;
	references: Reference[];
	devSrc: Source[];
	devDependencies: Dependency[];
	devReferences: Reference[];
	target?: Target;
	buildDir?: string;
}

/** Recognized keys in [project] / [package] sections. */
const KNOWN_SECTION_KEYS = new Set([
	"name",
	"version",
	"authors",
	"publish",
	"target",
	"build-dir",
	"codename"
]);

/** Snake_case → kebab-case corrections for common misspellings. */
const SNAKE_TO_KEBAB: Record<string, string> = {
	build_dir: "build-dir",
	src_encoding: "src-encoding"
};

function validateSectionKeys(metadata: Metadata, _section: string): void {
	const unknown = Object.keys(metadata);
	if (!unknown.length) return;

	const suggestions: string[] = [];

	for (const key of unknown) {
		// Direct match in the snake_case → kebab-case map
		if (SNAKE_TO_KEBAB[key]) {
			suggestions.push(`  "${key}" is not a valid key. Did you mean "${SNAKE_TO_KEBAB[key]}"?`);
			continue;
		}

		// Heuristic: any snake_case key that matches a known key after _ → -
		const kebab = key.replace(/_/g, "-");
		if (kebab !== key && KNOWN_SECTION_KEYS.has(kebab)) {
			suggestions.push(`  "${key}" is not a valid key. Did you mean "${kebab}"?`);
		}
	}

	if (suggestions.length) {
		manifestOk(false, suggestions.join("\n"));
	}
}

const EXAMPLE = `Example vbaproject.toml for a package (e.g. library to be shared):

  [package]
  name = "my-package"
  version = "0.0.0"
  authors = ["..."]

Example vbaproject.toml for a project (e.g. workbook, document, etc.):

  [project]
  name = "my-project"
  target = "xlsm"`;

export function parseManifest(value: any, dir: string): Manifest {
	manifestOk(
		value && (value.package || value.project),
		`A [package] or [project] section is required. \n\n${EXAMPLE}`
	);

	let type: ManifestType;
	let name: string;
	let version: string;
	let authors: string[] | undefined;
	let publish: boolean | undefined;
	let target: Target | undefined;
	let srcEncoding: string | undefined;
	let buildDir: string | undefined;
	let codename: string | undefined;
	let sectionMetadata: Metadata = {};

	if (value.project) {
		const {
			name: projectName,
			version: projectVersion,
			authors: projectAuthors,
			publish: projectPublish,
			target: projectTarget,
			"src-encoding": projectSrcEncoding,
			"build-dir": projectBuildDir,
			codename: projectCodename,
			...projectMetadata
		} = value.project;

		// src-encoding now lives under [source], not [project]
		if (projectSrcEncoding !== undefined) {
			manifestOk(
				false,
				`"src-encoding" should be set in the [source] section, not in [project].` +
					`\n\nMove it to [source]:\n\n  [source]\n  encoding = "${projectSrcEncoding}"`
			);
		}

		type = "project";
		name = projectName;
		version = projectVersion || DEFAULT_VERSION;
		authors = projectAuthors;
		publish = projectPublish;
		codename = projectCodename;
		sectionMetadata = projectMetadata;

		if (codename != null) {
			manifestOk(
				typeof codename === "string",
				`[project] codename must be a string (got ${typeof codename}).`
			);
		}

		manifestOk(name, `[project] name is a required field. \n\n${EXAMPLE}`);
		manifestOk(value.project.target, `[project] target is a required field. \n\n${EXAMPLE}`);

		target = parseTarget(projectTarget, name, dir);
		buildDir = projectBuildDir;
	} else {
		const {
			name: packageName,
			version: packageVersion,
			authors: packageAuthors,
			publish: packagePublish,
			target: packageTarget,
			"src-encoding": packageSrcEncoding,
			"build-dir": packageBuildDir,
			...packageMetadata
		} = value.package;

		// src-encoding now lives under [source], not [package]
		if (packageSrcEncoding !== undefined) {
			manifestOk(
				false,
				`"src-encoding" should be set in the [source] section, not in [package].` +
					`\n\nMove it to [source]:\n\n  [source]\n  encoding = "${packageSrcEncoding}"`
			);
		}

		type = "package";
		name = packageName;
		version = packageVersion;
		authors = packageAuthors;
		publish = packagePublish;
		sectionMetadata = packageMetadata;

		manifestOk(name, `[package] name is a required field. \n\n${EXAMPLE}`);
		manifestOk(version, `[package] version is a required field. \n\n${EXAMPLE}`);
		manifestOk(authors, `[package] authors is a required field. \n\n${EXAMPLE}`);

		target = packageTarget && parseTarget(packageTarget, name, dir);
		buildDir = packageBuildDir;
	}

	validateSectionKeys(sectionMetadata, type);

	// Normalize build-dir: strip trailing slashes, default to undefined
	if (typeof buildDir === "string") {
		buildDir = normalize(buildDir);
	}

	const src = parseSrc(value.src || {}, dir);
	const srcProperties = parseSrcProperties(value["source"]);
	srcEncoding = srcProperties?.encoding;
	const srcStructure = detectSrcStructure(src);
	const dependencies = parseDependencies(value.dependencies || {}, dir);
	const references = parseReferences(value.references || {});

	const devSrc = parseSrc(value["dev-src"] || {}, dir);
	const devDependencies = parseDependencies(value["dev-dependencies"] || {}, dir);
	const devReferences = parseReferences(value["dev-references"] || {});

	return {
		type,
		name,
		version,
		metadata: { authors, publish, ...sectionMetadata },
		src,
		srcEncoding,
		srcProperties,
		srcStructure,
		codename,
		dependencies,
		references,
		devSrc,
		devDependencies,
		devReferences,
		target,
		buildDir
	};
}

export async function loadManifest(dir: string): Promise<Manifest> {
	let file = join(dir, "vbaproject.toml");

	if (!(await pathExists(file))) {
		// Fallback to legacy manifest name for backward compatibility
		// (e.g. packages downloaded from registry that still use vba-block.toml)
		const legacyFile = join(dir, "vba-block.toml");
		if (await pathExists(legacyFile)) {
			file = legacyFile;
		} else {
			throw new CliError(
				ErrorCode.ManifestNotFound,
				dedent`
          vbaproject.toml not found in "${dir}".

          Try "vbapm init" to start a new project in this directory
          or "cd YOUR_PROJECTS_DIRECTORY" to change to a folder that contains an existing project.
        `
			);
		}
	}

	const raw = await readFile(file);

	let parsed;
	try {
		parsed = await parseToml(raw.toString());
	} catch (err: any) {
		throw new CliError(
			ErrorCode.ManifestInvalid,
			dedent`
				vbaproject.toml is invalid:

				Syntax Error: ${file} (${err?.line}:${err?.column})\n\n${err?.message || err}
			`
		);
	}

	const manifest = parseManifest(parsed, normalize(dir));

	return manifest;
}

export async function writeManifest(manifest: Manifest, dir: string) {
	const value = formatManifest(manifest, dir);
	const path = join(dir, "vbaproject.toml");
	let toml: string;

	if (await pathExists(path)) {
		const existing = await readFile(path, "utf8");
		toml = await patchToml(existing, value);
	} else {
		toml = await convertToToml(value);
	}

	await writeFile(path, toml);
}

export function formatManifest(manifest: Manifest, dir: string): object {
	const {
		type,
		name,
		version,
		metadata: { authors, publish, ...metadata }
	} = manifest;

	const values: any = { name };
	if (version !== DEFAULT_VERSION) values.version = version;
	if (authors != null) values.authors = authors;
	if (publish != null) values.publish = publish;
	Object.assign(values, metadata);

	const value: any = {
		[type]: values
	};

	if (manifest.target) {
		value[type].target = formatTarget(manifest.target, manifest.name, dir);
	}

	if (manifest.buildDir != null && manifest.buildDir !== "build") {
		values["build-dir"] = manifest.buildDir;
	}

	if (manifest.codename != null && manifest.codename !== "VBAProject") {
		values["codename"] = manifest.codename;
	}

	if (manifest.srcProperties) {
		value["source"] = manifest.srcProperties;
	}

	value.src = formatSrc(manifest.src, dir);

	if (manifest.dependencies.length) {
		value.dependencies = formatDependencies(manifest.dependencies, dir);
	}
	if (manifest.references.length) {
		value.references = formatReferences(manifest.references);
	}

	if (manifest.devSrc.length) {
		value["dev-src"] = formatSrc(manifest.devSrc, dir);
	}
	if (manifest.devDependencies.length) {
		value["dev-dependencies"] = formatDependencies(manifest.devDependencies, dir);
	}
	if (manifest.devReferences.length) {
		value["dev-references"] = formatReferences(manifest.devReferences);
	}

	return value;
}
