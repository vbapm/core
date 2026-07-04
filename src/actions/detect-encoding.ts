import { readFile } from "../utils/fs";
import {
	codepageToLabel,
	getSystemCodepage,
	SUPPORTED_WINDOWS_CODEPAGE_LABELS
} from "../build/encoding-sniffer";

/**
 * Decide which encoding to declare for a newly imported VBA project.
 *
 * a higher-level policy decision.
 *
 * That policy is:
 * 1. Default to the system codepage.
 * 2. If jschardet confidently (≥ 80 %) detects a *different* Windows
 *    codepage, adopt that instead (e.g. a Japanese workbook opened on
 *    a Western European machine).

 * ## Relationship to encoding-sniffer
 *
 * {@link ../build/encoding-sniffer.ts | encoding-sniffer} answers *what*
 * encoding a buffer is (BOM detection, UTF-8 heuristic, decode). This
 * module answers *which* encoding the user should put in their
 * `vbaproject.toml` — a higher-level policy decision.
 *
 * **Separation of concerns** — encoding-sniffer stays focused on
 * low-level byte detection; this module owns the business rule that
 * combines system state (ACP), file contents, and jschardet
 * confidence into a single label the caller writes to TOML.
 *
 * @param firstSourcePath - Absolute path to the first source file (used
 *   as a sample for jschardet detection).
 * @returns The detected encoding label (e.g. "cp1252", "cp932"), or
 *   `undefined` if all files are ASCII-only.
 */
export async function detectImportEncoding(firstSourcePath: string): Promise<string | undefined> {
	// Read the file to check for non-ASCII and run detection
	let buffer: Buffer;
	try {
		buffer = await readFile(firstSourcePath);
	} catch {
		return undefined;
	}

	// Check for non-ASCII characters
	const code = buffer.toString();
	if (!/[^\x00-\x7F]/.test(code)) return undefined;

	// Default to system codepage
	const systemCp = getSystemCodepage();
	const systemLabel = codepageToLabel(systemCp);
	let detectedEncoding = systemLabel;

	// Try jschardet to see if a different encoding is more likely
	try {
		// TODO: Fix when we update vbapm to ESM-only.
		const jschardet = (await import("jschardet")).default;
		if (!jschardet) {
			throw new Error("jschardet not available");
		}

		const results = (jschardet.detectAll(buffer) as Array<{ encoding: string; confidence: number }>)
			.filter((r: { encoding: string }) => SUPPORTED_WINDOWS_CODEPAGE_LABELS.has(r.encoding))
			.sort((a: { confidence: number }, b: { confidence: number }) => b.confidence - a.confidence);

		const SYSTEM_PREFIX = /^windows-?/i;
		if (results.length > 0 && results[0].confidence >= 0.4) {
			const detected = results[0].encoding.toLowerCase();
			// jschardet may return "SHIFT_JIS" for Windows-932 content;
			// normalize to "cp932" which is our canonical label.
			const normalized = detected === "shift_jis" ? "cp932" : detected;
			const systemNormalized = systemLabel.replace(SYSTEM_PREFIX, "cp").toLowerCase();

			if (normalized !== systemNormalized) {
				detectedEncoding = normalized;
			}
		}
	} catch {
		// Log error message to console as a warning, but don't fail the import. The system codepage is a safe fallback.
		console.warn("[detectImportEncoding] jschardet unavailable or failed; using system codepage:", systemLabel);
	}

	return detectedEncoding;
}
