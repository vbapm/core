import dedent from "@timhall/dedent";
import { manifestOk } from "../errors";
import { isString } from "../utils/is";
import { join, relative } from "../utils/path";

/*
  # Source

  A source file to be imported

  source: path | { path }
*/

export interface Source {
	name: string;
	path: string;
	binary?: string;
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
	const { path: relativePath, binary } = value;

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

	return source;
}

export function formatSrc(src: Source[], dir: string): object {
	const value: { [name: string]: string } = {};
	src.forEach(source => {
		let { name, path } = source;
		path = relative(dir, path);
		value[name] = path;
	});

	return value;
}
