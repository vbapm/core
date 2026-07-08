import { extname } from "../utils/path";
import { Source } from "./source";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** Maps VBA component types to subdirectories under `src/`. */
export interface SrcSubfolders {
	Modules?: string;
	Forms?: string;
	Classes?: string;
	Objects?: string;
}

/** Parsed form of the optional `[src-properties]` TOML section. */
export interface SrcProperties {
	sort?: {
		"by-types"?: boolean;
		alphabetical?: boolean;
	};
	subfolders?: SrcSubfolders;
}

/**
 * Resolve the subdirectory under `src/` for a given component type.
 * Uses the `subfolders` config from `[src-properties]` if present,
 * otherwise defaults to placing all files directly in `src/`.
 */
export function resolveSrcSubfolders(subfolders: SrcSubfolders | undefined, type: string): string {
	if (!subfolders) return "";

	const key =
		type === "object"
			? "Objects"
			: type === "class"
				? "Classes"
				: type === "form"
					? "Forms"
					: "Modules";
	return subfolders[key] || "";
}

/** Describes how the `[src]` section is currently organised. */
export interface SrcStructure {
	/** true when all .bas files are contiguous, all .frm contiguous, all .cls contiguous. */
	sortedByTypes: boolean;

	/** true when all entries are in alphabetical order, globally. */
	sortedAlphabetically: boolean;

	/** Convenience: true when both sortedByTypes AND within-each-type alphabetical. */
	sortedByTypeThenAlphabetically: boolean;

	/** true when none of the above patterns are detected. */
	unstructured: boolean;
}

// ---------------------------------------------------------------------------
// Detection
// ---------------------------------------------------------------------------

/**
 * Detect the organisational structure of the `[src]` section from the parsed
 * source entries alone — no `[src-properties]` needed.
 */
export function detectSrcStructure(src: Source[]): SrcStructure {
	// ---- empty src ----
	if (src.length === 0) {
		return {
			sortedByTypes: false,
			sortedAlphabetically: false,
			sortedByTypeThenAlphabetically: false,
			unstructured: true
		};
	}

	// ---- sortedByTypes check ----
	let sortedByTypes = true;
	let currentExt = "";
	let segmentCount = 0;
	const seenExts = new Set<string>();
	for (const s of src) {
		const ext = extname(s.path).toLowerCase();
		if (ext !== currentExt) {
			// Same extension appearing again after a different type → not contiguous
			if (seenExts.has(ext)) {
				sortedByTypes = false;
				break;
			}
			seenExts.add(currentExt);
			segmentCount++;
			currentExt = ext;
		}
	}
	if (sortedByTypes && segmentCount > 4) sortedByTypes = false;

	// ---- sortedAlphabetically check ----
	let sortedAlphabetically = true;
	for (let i = 1; i < src.length; i++) {
		if (src[i].name.toLowerCase() < src[i - 1].name.toLowerCase()) {
			sortedAlphabetically = false;
			break;
		}
	}

	// ---- sortedByTypeThenAlphabetically ----
	let sortedByTypeThenAlphabetically = false;
	if (sortedByTypes) {
		sortedByTypeThenAlphabetically = true;
		// verify alphabetical within each segment
		let segStart = 0;
		let segExt = extname(src[0].path).toLowerCase();
		for (let i = 1; i <= src.length; i++) {
			const ext = i < src.length ? extname(src[i].path).toLowerCase() : "";
			if (ext !== segExt || i === src.length) {
				// check this segment [segStart, i-1] is alphabetical
				for (let j = segStart + 1; j < i; j++) {
					if (src[j].name.toLowerCase() < src[j - 1].name.toLowerCase()) {
						sortedByTypeThenAlphabetically = false;
						break;
					}
				}
				if (!sortedByTypeThenAlphabetically) break;
				segStart = i;
				segExt = ext;
			}
		}
	}

	// ---- unstructured ----
	const unstructured = !(sortedByTypes || sortedAlphabetically);

	return {
		sortedByTypes,
		sortedAlphabetically,
		sortedByTypeThenAlphabetically,
		unstructured
	};
}

/**
 * Parse the `[src-properties]` TOML section into a typed object.
 * Returns `undefined` when the section is absent.
 */
export function parseSrcProperties(raw: any): SrcProperties | undefined {
	if (!raw || typeof raw !== "object") return undefined;

	const props: SrcProperties = {};

	if (raw.sort && typeof raw.sort === "object") {
		const sort: SrcProperties["sort"] = {};
		if (typeof raw.sort["by-types"] === "boolean") sort["by-types"] = raw.sort["by-types"];
		if (typeof raw.sort.alphabetical === "boolean") sort.alphabetical = raw.sort.alphabetical;
		if (Object.keys(sort).length > 0) props.sort = sort;
	}

	if (raw.subfolders && typeof raw.subfolders === "object") {
		const sf: SrcSubfolders = {};
		if (typeof raw.subfolders.Modules === "string") sf.Modules = raw.subfolders.Modules;
		if (typeof raw.subfolders.Forms === "string") sf.Forms = raw.subfolders.Forms;
		if (typeof raw.subfolders.Classes === "string") sf.Classes = raw.subfolders.Classes;
		if (typeof raw.subfolders.Objects === "string") sf.Objects = raw.subfolders.Objects;
		if (Object.keys(sf).length > 0) props.subfolders = sf;
	}

	return Object.keys(props).length > 0 ? props : undefined;
}

/**
 * Post-process a TOML string to insert blank lines between type groups
 * in the `[src]` section. Only used for initial creation of individual-listing
 * manifests (`vba init --list-all`).
 */
export function insertTypeGroupBlankLines(toml: string): string {
	const lines = toml.split("\n");
	const result: string[] = [];
	let inSrc = false;
	let lastExt = "";

	for (const line of lines) {
		// Detect entering [src] section
		if (/^\[src\]/.test(line.trim())) {
			inSrc = true;
			result.push(line);
			continue;
		}

		// Detect leaving [src] section (next section header or end)
		if (inSrc && /^\[/.test(line.trim())) {
			inSrc = false;
		}

		if (inSrc && line.includes("=")) {
			const extMatch = line.match(/\.(bas|frm|cls)\b/i);
			const ext = extMatch ? extMatch[1].toLowerCase() : "";
			if (ext && lastExt && ext !== lastExt) {
				result.push(""); // blank line between type groups
			}
			if (ext) lastExt = ext;
		}

		result.push(line);
	}

	return result.join("\n");
}
