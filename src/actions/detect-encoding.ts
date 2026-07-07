import { readFile } from "../utils/fs";
import {
	Codepage,
	codepageToLabel,
	getSystemCodepage,
	labelToCodepage,
	SUPPORTED_WINDOWS_CODEPAGE_LABELS
} from "../build/encoding-sniffer";

import jschardet from "jschardet";

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

	// Check for non-ASCII characters — byte loop avoids allocating a string
	let hasNonAscii = false;
	for (let i = 0; i < buffer.length; i++) {
		if (buffer[i] > 0x7f) {
			hasNonAscii = true;
			break;
		}
	}
	if (!hasNonAscii) return undefined;

	// Default to system codepage
	const systemCp = getSystemCodepage();
	let detectedEncoding = codepageToLabel(systemCp);

	// Try jschardet to see if a different encoding is more likely
	try {
		const results = (
			jschardet.detectAll(buffer) as Array<{ encoding: string; confidence: number }>
		).sort((a: { confidence: number }, b: { confidence: number }) => b.confidence - a.confidence);

		// Remap jschardet guesses that are impossible in a VBA context
		// (e.g. MacCyrillic will never be a VBA source encoding)
		const remapped = results.map(r => ({
			...r,
			encoding: remapForVbaContext(r.encoding)
		}));

		const filtered = remapped.filter((r: { encoding: string }) =>
			SUPPORTED_WINDOWS_CODEPAGE_LABELS.has(r.encoding)
		);

		if (filtered.length > 0 && filtered[0].confidence >= 0.4) {
			const detectedLabel = filtered[0].encoding.toLowerCase();
			const detectedCp = labelToCodepage(detectedLabel);

			if (detectedCp !== Codepage.Unknown && detectedCp !== systemCp) {
				// Normalize to canonical label (e.g. "shift_jis" → "windows-932")
				detectedEncoding = codepageToLabel(detectedCp);
			}
		}
	} catch {
		// jschardet unavailable — keep system codepage
	}

	return detectedEncoding;
}

/**
 * Remap jschardet encoding guesses that are impossible or extremely
 * unlikely in a VBA context to their Windows codepage equivalents.
 *
 * VBA on Windows only uses Windows ANSI / DBCS codepages. Encodings
 * like MacCyrillic or ISO-8859-* will never appear in a VBA source
 * file exported from Excel, but jschardet may guess them based on
 * byte frequency. We remap them to the corresponding Windows codepage.
 */
function remapForVbaContext(encoding: string): string {
	// SHIFT_JIS → Windows-932 (VBA uses Windows-31J, not standard Shift_JIS)
	if (/^shift[-_]?jis$/i.test(encoding)) return "windows-932";
	// MacCyrillic → Windows-1251 (VBA on Windows never uses Mac encodings)
	if (/^x?-?mac-?cyrillic$/i.test(encoding)) return "windows-1251";
	// ISO-8859-5 (Cyrillic) → Windows-1251
	if (/^iso-?8859-?5$/i.test(encoding)) return "windows-1251";
	// ISO-8859-2 (Central European) → Windows-1250
	if (/^iso-?8859-?2$/i.test(encoding)) return "windows-1250";

	return encoding;
}
