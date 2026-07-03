import { manifestOk } from "../errors";
import { has } from "../utils/has";
import { isString } from "../utils/is";
import { join, relative, sanitize } from "../utils/path";

/*
  # Target
  # Indicates what format is targeted for the build ouput and where the input 
  # source files are located.

  target: type | { type, name?, path? }

  By default "target" is used for path
*/

export type TargetType = "xlsx" | "xlsm" | "xlam";

export interface Target {
	/**
	 * The name of the target file (without extension). If not specified, the package name will be used.
	 */
	name: string;
	/**
	 * The type of the target file. Must be one of "xlsx", "xlsm", or "xlam".
	 */
	type: TargetType;
	/**
	 * The path to directory with the input files used to build the target. If not specified, the default path "target" will be used.
	 */
	path: string;
	/**
	 * The filename of the target file (with extension). This is derived from the name and type.
	 */
	filename: string;
	/** Encoding for the target (e.g. "windows-1252", "cp932"). Defaults to system codepage. */
	encoding?: string;
}

const TARGET_TYPES = ["xlsx", "xlsm", "xlam"];

const EXAMPLE = `Example vbaproject.toml:

  [project]
  target = "xlsm"

Example vbaproject.toml with alternative path:

  [project]
  target = { type = "xlsm", path = "target/xlsm" }`;

export function parseTarget(value: any, pkgName: string, dir: string): Target {
	if (isString(value)) value = { type: value };
	if (!has(value, "name")) value = { name: pkgName, ...value };
	const { type, name, path: relativePath = "target", encoding } = value;

	manifestOk(isString(type), `Target is missing <type>. \n\n${EXAMPLE}.`);
	manifestOk(
		isSupportedTargetType(type),
		`Unsupported target type <${type}>. Only <xlsx>, <xlsm>, and <xlam> are supported currently.`
	);

	const path = join(dir, relativePath);
	const filename = `${sanitize(name)}.${type}`;

	const target: Target = { name, type, path, filename };
	if (encoding) target.encoding = encoding;

	return target;
}

export function isSupportedTargetType(type: string): type is TargetType {
	return isString(type) && TARGET_TYPES.includes(type);
}

export function formatTarget(target: Target, defaultName: string, dir: string): string | object {
	let { name, type: targetType, path, encoding } = target;
	path = relative(dir, path);

	let value: string | { type: string; name?: string; path?: string; encoding?: string };
	if (name !== defaultName || path !== "target" || encoding) {
		value = { type: targetType };
		if (name !== defaultName) value.name = name;
		if (path !== "target") value.path = path;
		if (encoding) value.encoding = encoding;
	} else {
		value = targetType;
	}

	return value;
}
