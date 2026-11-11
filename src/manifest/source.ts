import { Minimatch } from "minimatch";
import dedent from "@timhall/dedent";
import { manifestOk } from "../errors";
import { isString } from "../utils/is";
import { extname, join, normalize, relative } from "../utils/path";

/*
  # Source

  A source file to be imported

  source: path | { path }
*/

export interface Source {
	name: string;
	path: string;
	binary?: string;
	encoding?: string;
}

export function isPowerQuerySource(source: Source): boolean {
	return isPowerQueryPath(source.path);
}

export function isPowerQueryPath(path: string): boolean {
	return extname(path).toLowerCase() === ".pq";
}

export function resolveSourceEncoding(
	sources: Source[],
	filePath: string,
	projectDir: string,
	defaultEncoding?: string
): string | undefined {
	const matchingSources = sources.filter(source =>
		sourcePathMatches(source.path, filePath, projectDir)
	);
	const explicitEncoding = matchingSources.find(
		source => !source.path.includes("*") && source.encoding
	)?.encoding;
	if (explicitEncoding) return explicitEncoding;

	const wildcardEncoding = matchingSources.find(
		source => source.path.includes("*") && source.encoding
	)?.encoding;
	return wildcardEncoding ?? defaultEncoding;
}

export function sourcePathMatches(
	sourcePath: string,
	filePath: string,
	projectDir: string
): boolean {
	const pattern = relativeToProject(sourcePath, projectDir);
	const normalisedPath = relativeToProject(filePath, projectDir);

	if (!sourcePath.includes("*")) return pattern === normalisedPath;

	return new Minimatch(pattern).match(normalisedPath);
}

function relativeToProject(path: string, projectDir: string): string {
	const normalizedPath = path.replace(/\\/g, "/");
	const normalizedProjectDir = projectDir.replace(/\\/g, "/").replace(/\/+$/, "");
	const relativePath =
		normalizedPath === normalizedProjectDir || normalizedPath.startsWith(`${normalizedProjectDir}/`)
			? relative(projectDir, path)
			: path;

	return normalize(relativePath).replace(/\\/g, "/").replace(/^\.\//, "");
}

const EXAMPLE = `Example vbaproject.toml:

  [source.files]
  A = "src/a.bas"
  B = { path = "src/b.cls" }`;

const VBA_EXTENSIONS = /^(bas|cls|frm|doccls)$/i;

export function parseSrc(value: any, dir: string): Source[] {
	return Object.entries(value).map(([name, value]) => parseSource(name, value, dir));
}

export function parseSource(name: string, value: string | any, dir: string): Source {
	if (isString(value)) value = { path: value };
	const { path: relativePath, binary, encoding } = value;

	if (!relativePath) {
		const extKey = Object.keys(value).find(k => VBA_EXTENSIONS.test(k));
		if (extKey) {
			manifestOk(
				false,
				dedent`
				src key <${name}.${extKey}> should not include the file extension.
				You want to use a <${name}> instead.
				`
			);
		}
	}

	manifestOk(relativePath, `src <${name}> is missing path. \n\n${EXAMPLE}`);
	const path = join(dir, relativePath);

	const source: Source = { name, path };
	if (binary) source.binary = join(dir, binary);
	if (encoding) source.encoding = encoding;

	return source;
}

export function formatSrc(src: Source[], dir: string): object {
	const value: {
		[name: string]: string | { path: string; binary?: string; encoding?: string };
	} = {};
	src.forEach(source => {
		const formatPath = (path: string) => {
			// Only relativize absolute paths; wildcard entries are already relative
			if (path.match(/^[a-zA-Z]:[\\/]/) || path.startsWith("/")) {
				path = relative(dir, path);
			}
			// Standardize to forward slashes even on Windows
			return normalize(path);
		};

		const path = formatPath(source.path);
		if (source.encoding || source.binary) {
			value[source.name] = {
				path,
				...(source.binary ? { binary: formatPath(source.binary) } : {}),
				...(source.encoding ? { encoding: source.encoding } : {})
			};
		} else {
			value[source.name] = path;
		}
	});

	return value;
}
