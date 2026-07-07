import { extname } from "../utils/path";
import { Source } from "./source";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** Parsed form of the optional `[src-properties]` TOML section. */
export interface SrcProperties {
	grouping?: boolean;
	sort?: {
		"by-types"?: boolean;
		alphabetical?: boolean;
	};
}

/** Describes how the `[src]` section is currently organised. */
export interface SrcStructure {
	/** true when [src] uses the 3 reserved grouped keys (Modules, Forms, Classes). */
	grouped: boolean;

	/** true when all .bas files are contiguous, all .frm contiguous, all .cls contiguous. */
	sortedByTypes: boolean;

	/** true when all entries are in alphabetical order, globally. */
	sortedAlphabetically: boolean;

	/** Convenience: true when both sortedByTypes AND within-each-type alphabetical. */
	sortedByTypeThenAlphabetically: boolean;

	/** true when none of the above patterns are detected. */
	unstructured: boolean;

	/** When grouped, the raw glob strings keyed by type. */
	groupedPatterns?: Record<"Modules" | "Forms" | "Classes", string | string[]>;
}

// ---------------------------------------------------------------------------
// Detection
// ---------------------------------------------------------------------------

const GROUPED_KEYS = new Set(["modules", "forms", "classes"]);

/**
 * Detect the organisational structure of the `[src]` section from the parsed
 * source entries alone — no `[src-properties]` needed.
 */
export function detectSrcStructure(src: Source[]): SrcStructure {
	// ---- grouped check ----
	if (
		src.length === 3 &&
		GROUPED_KEYS.has(src[0].name.toLowerCase()) &&
		GROUPED_KEYS.has(src[1].name.toLowerCase()) &&
		GROUPED_KEYS.has(src[2].name.toLowerCase())
	) {
		const groupedPatterns: Record<string, string | string[]> = {};
		for (const s of src) {
			groupedPatterns[s.name] = s.path;
		}
		return {
			grouped: true,
			sortedByTypes: false,
			sortedAlphabetically: false,
			sortedByTypeThenAlphabetically: false,
			unstructured: false,
			groupedPatterns: groupedPatterns as SrcStructure["groupedPatterns"]
		};
	}

	// ---- sortedByTypes check ----
	let sortedByTypes = true;
	let currentExt = "";
	let segmentCount = 0;
	for (const s of src) {
		const ext = extname(s.path).toLowerCase();
		if (ext !== currentExt) {
			segmentCount++;
			currentExt = ext;
		}
	}
	if (segmentCount > 3) sortedByTypes = false;

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
	if (sortedByTypes && sortedAlphabetically) {
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
		grouped: false,
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

	if (typeof raw.grouping === "boolean") props.grouping = raw.grouping;

	if (raw.sort && typeof raw.sort === "object") {
		const sort: SrcProperties["sort"] = {};
		if (typeof raw.sort["by-types"] === "boolean") sort["by-types"] = raw.sort["by-types"];
		if (typeof raw.sort.alphabetical === "boolean") sort.alphabetical = raw.sort.alphabetical;
		if (Object.keys(sort).length > 0) props.sort = sort;
	}

	return Object.keys(props).length > 0 ? props : undefined;
}

/**
 * Post-process a TOML string to insert blank lines between type groups
 * in the `[src]` section. Only used for initial creation of individual-listing
 * manifests (`vba init --individual`).
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
