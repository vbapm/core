import { manifestOk } from "../errors";
import { relative, resolve } from "../utils/path";

/*
  # Reference

  Two kinds of references are supported:

  ## COM/DLL reference

  To add reference, guid, major version, and minor version are required

  reference: { guid, version }

  Where version is "MAJOR.MINOR"

  ## Peer reference (VBA project-to-project)

  References to another VBA project (.xlam/.xlsm) have no GUID or version.
  They are declared with `peer = true` and optionally a `path`:

  reference: { peer = true, path? }

  The path is stored relative when the peer lives inside the project folder or a
  sibling folder, otherwise as an absolute path.
*/

export interface ReferenceDetails {
	dependency?: string;
}

export interface Reference {
	name: string;
	guid: string; // "" for peer references
	major: number; // 0 for peer references
	minor: number; // 0 for peer references
	/** True for references to another VBA project (addin/workbook). */
	peer?: boolean;
	/** Stored path for peer references (relative if nearby, else absolute). */
	path?: string;
}

const VERSION_REGEX = /^(\d+)\.(\d+)$/;
const GUID_REGEX = /\{.{8}-.{4}-.{4}-.{4}-.{12}\}/;

const isVersion = (value: string) => VERSION_REGEX.test(value);
const isGuid = (value: string) => GUID_REGEX.test(value);

const toInt = (value: string) => parseInt(value, 10);
const getMajorMinor = (version: string) => {
	const [major, minor] = version.split(".", 2).map(toInt);

	return { major, minor };
};

const EXAMPLE = `Example vbaproject.toml:

  [references.Scripting]
  version = "1.0"
  guid = "{420B2830-E718-11CF-893D-00A0C9054228}"`;

/** Absolute path detector for normalized (forward-slash) paths. */
const ABSOLUTE_REGEX = /^([a-zA-Z]:)?\//;

/**
 * Choose the path form to store for a peer reference.
 *
 * If the peer file lives inside the project folder (no ".." segments) or inside a
 * sibling folder (a single leading ".." segment), keep it relative for portability.
 * Otherwise fall back to the absolute path — the user can edit it to a relative one.
 */
export function relativizePeerPath(fromDir: string, filePath: string): string {
	if (!ABSOLUTE_REGEX.test(filePath)) return filePath;

	const rel = relative(fromDir, filePath);
	if (ABSOLUTE_REGEX.test(rel)) return filePath;

	const upCount = rel.split("/").filter(segment => segment === "..").length;
	return upCount <= 1 ? rel : filePath;
}

/**
 * Resolve peer reference paths to absolute form for import.
 *
 * The manifest stores peer paths relative (when nearby) or absolute. Before
 * passing them to the VBA addin (`References.AddFromFile` needs an absolute
 * path), resolve relative paths against the project folder. References without
 * a path are passed through unchanged — validation is the caller's job.
 */
export function resolvePeerReferencePaths(references: Reference[], dir: string): Reference[] {
	return references.map(reference => {
		if (!reference.peer || !reference.path) return reference;
		return { ...reference, path: resolve(dir, reference.path) };
	});
}

export function parseReferences(value: any): Reference[] {
	return Object.entries(value).map(([name, value]) => parseReference(name, value));
}

export function parseReference(name: string, value: any): Reference {
	const { version, guid, peer, path } = value;

	// Peer reference to another VBA project — no GUID or version
	if (peer) {
		return { name, guid: "", major: 0, minor: 0, peer: true, path };
	}

	manifestOk(
		isVersion(version),
		`Reference <${name}> has an invalid version <${version}>. \n\n${EXAMPLE}.`
	);
	manifestOk(isGuid(guid), `Reference <${name}> has an invalid guid <${guid}>. \n\n${EXAMPLE}.`);

	const { major, minor } = getMajorMinor(version);

	return { name, guid, major, minor };
}

export function formatReferences(references: Reference[]): object {
	const value: { [name: string]: object } = {};
	references.forEach(reference => {
		const { name, guid, major, minor, peer, path } = reference;

		if (peer) {
			const entry: { [key: string]: any } = { peer: true };
			if (path) entry.path = path;
			value[name] = entry;
		} else {
			const version = `${major}.${minor}`;
			value[name] = { version, guid };
		}
	});

	return value;
}
