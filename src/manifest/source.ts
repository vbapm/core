import dedent from "@timhall/dedent";
import { manifestOk } from "../errors";
import { isString } from "../utils/is";
import { join, relative } from "../utils/path";
import { SrcStructure } from "./src-sort";

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

const EXAMPLE = `Example vbaproject.toml:

  [src]
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

export function formatSrc(src: Source[], dir: string, structure?: SrcStructure): object {
	// Grouped convention: emit the 3 reserved keys
	if (structure?.grouped) {
		const value: Record<string, string | string[]> = {};
		for (const s of src) {
			const key = s.name; // "Modules", "Forms", or "Classes"
			const rawPath = s.path;
			// If the path is absolute, make it relative
			const relPath = rawPath.startsWith(dir) ? relative(dir, rawPath) : rawPath;
			// If the value is an array of globs (from groupedPatterns), preserve it
			if (structure.groupedPatterns?.[key as "Modules" | "Forms" | "Classes"]) {
				const patterns = structure.groupedPatterns[key as "Modules" | "Forms" | "Classes"];
				value[key] = typeof patterns === "string" ? patterns : patterns;
			} else {
				value[key] = relPath;
			}
		}
		return value;
	}

	// Individual convention: emit entries in array order
	const value: { [name: string]: string } = {};
	src.forEach(source => {
		let { name, path } = source;
		path = relative(dir, path);
		value[name] = path;
	});

	return value;
}
