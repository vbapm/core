/**
 * Encoding detection for VBA source files.
 *
 * VBA's `Component.Export` writes files in the system's ANSI codepage
 * (e.g., Windows-1252 on Western Windows). This sniffer detects the
 * actual encoding of a buffer so that collaboration across machines
 * with different locale codepages works correctly.
 *
 * Detection order:
 * 1. UTF-16 BOM (LE/BE)
 * 2. UTF-8 BOM
 * 3. UTF-16 without BOM (null-byte heuristic)
 * 4. Validate as UTF-8
 * 5. Fall back to windows-1252 (most common ANSI codepage)
 */

import * as iconv from "iconv-lite";

/**
 * Known codepage for a VBA source buffer.
 *
 * Use {@link Codepage.Unknown} when the encoding is not known ahead of
 * time — the sniffer will auto-detect it. Use a specific value (e.g.
 * {@link Codepage.Windows1252}) when the caller knows the encoding (e.g.
 * files just exported by VBA's `Component.Export`).
 *
 * Codepage numbers match the Windows code page identifiers:
 * https://learn.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
 */
export enum Codepage {
	Unknown = 0,

	// ANSI codepages (used by VBA's Component.Export on different locales)
	Windows1250 = 1250, // Central European
	Windows1251 = 1251, // Cyrillic
	Windows1252 = 1252, // Western European (includes French)
	Windows1253 = 1253, // Greek
	Windows1254 = 1254, // Turkish
	Windows1255 = 1255, // Hebrew
	Windows1256 = 1256, // Arabic
	Windows1257 = 1257, // Baltic
	Windows1258 = 1258, // Vietnamese

	Windows932 = 932, // Japanese (Windows-31J, not Shift_JIS — see codepageToLabel)
	Windows936 = 936, // Simplified Chinese (GBK)
	Windows874 = 874, // Thai
	Windows949 = 949, // Korean (Unified Hangul Code - ks_c_5601-1987)
	Windows950 = 950, // Traditional Chinese (Big5)

	UTF8 = 65001
}

export type SniffedEncoding = "utf8" | "utf16le" | "utf16be" | "windows-1252";

export interface SniffResult {
	encoding: SniffedEncoding;
	hasBom: boolean;
}

/**
 * Map a {@link Codepage} to a {@link TextDecoder} label.
 * Returns `""` for {@link Codepage.Unknown} (caller must sniff instead).
 */
export function codepageToLabel(codepage: Codepage): string {
	return CODEPAGE_LABELS[codepage] ?? "";
}

const CODEPAGE_LABELS: Record<number, string> = {
	[Codepage.UTF8]: "utf-8",
	[Codepage.Windows1250]: "windows-1250", // Central European
	[Codepage.Windows1251]: "windows-1251", // Cyrillic
	[Codepage.Windows1252]: "windows-1252", // Western European (includes English, Spanish, French, Portuguese, German except ẞ, Italian and more)
	[Codepage.Windows1253]: "windows-1253", // Greek
	[Codepage.Windows1254]: "windows-1254", // Turkish
	[Codepage.Windows1255]: "windows-1255", // Hebrew
	[Codepage.Windows1256]: "windows-1256", // Arabic
	[Codepage.Windows1257]: "windows-1257", // Baltic
	[Codepage.Windows1258]: "windows-1258", // Vietnamese
	// For Japanese, use "windows-932" (Windows-31J), not "shift_jis".
	// The IANA Shift_JIS differs from the Windows codepage.
	// Namely, 0x5C → ¥ and 0x7E → ‾, while the
	// Windows codepage maps them to \ and ~ respectively.
	// READ MORE: https://en.wikipedia.org/wiki/Code_page_932_(Microsoft_Windows)#Differences_from_standard_Shift_JIS
	[Codepage.Windows932]: "windows-932", // Japanese (Windows-31J, not Shift_JIS)
	[Codepage.Windows936]: "windows-936", // Simplified Chinese (GBK)
	[Codepage.Windows874]: "windows-874", // Thai
	[Codepage.Windows949]: "windows-949", // Korean (ks_c_5601-1987)
	[Codepage.Windows950]: "windows-950" // Traditional Chinese (Big5)
};

/**
 * Detect the encoding of a buffer containing VBA source code.
 * Unlike {@link ../encoding.ts}, this function validates UTF-8 and falls
 * back to a legacy codepage when the buffer is not valid UTF-8.
 */
export function sniffEncoding(buffer: Buffer): SniffResult {
	// ── BOM detection ──────────────────────────────────────────

	if (buffer.length >= 2) {
		if (buffer[0] === 0xff && buffer[1] === 0xfe) {
			return { encoding: "utf16le", hasBom: true };
		}
		if (buffer[0] === 0xfe && buffer[1] === 0xff) {
			return { encoding: "utf16be", hasBom: true };
		}
	}

	if (buffer.length >= 3 && buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf) {
		return { encoding: "utf8", hasBom: true };
	}

	// ── UTF-16 without BOM (null-byte heuristic) ───────────────

	let oddNulls = 0;
	let evenNulls = 0;
	for (let i = 0; i < buffer.length; i++) {
		if (buffer[i] !== 0x00) continue;
		if (i % 2 === 0) {
			evenNulls++;
		} else {
			oddNulls++;
		}
	}

	if (oddNulls > 8 && oddNulls > evenNulls * 2) {
		return { encoding: "utf16le", hasBom: false };
	}
	if (evenNulls > 8 && evenNulls > oddNulls * 2) {
		return { encoding: "utf16be", hasBom: false };
	}

	// ── UTF-8 validation ───────────────────────────────────────

	if (isValidUtf8(buffer)) {
		return { encoding: "utf8", hasBom: false };
	}

	// ── Fallback: legacy ANSI codepage ─────────────────────────

	return { encoding: "windows-1252", hasBom: false };
}

/** Known ACP → Codepage mapping for getSystemCodepage. */
const ACP_TO_CODEPAGE: Record<string, Codepage> = {
	"932": Codepage.Windows932,
	"874": Codepage.Windows874,
	"936": Codepage.Windows936,
	"949": Codepage.Windows949,
	"950": Codepage.Windows950,
	"1250": Codepage.Windows1250,
	"1251": Codepage.Windows1251,
	"1252": Codepage.Windows1252,
	"1253": Codepage.Windows1253,
	"1254": Codepage.Windows1254,
	"1255": Codepage.Windows1255,
	"1256": Codepage.Windows1256,
	"1257": Codepage.Windows1257,
	"1258": Codepage.Windows1258
};

/**
 * Return the local system ANSI codepage by reading the ACP value
 * from the Windows registry. Falls back to {@link Codepage.Windows1252}
 * if the registry cannot be read.
 */
export function getSystemCodepage(): Codepage {
	try {
		const { execSync } = require("child_process");
		const result = execSync(
			'reg query "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Nls\\CodePage" /v ACP',
			{ encoding: "utf8", timeout: 5000 }
		);
		const match = result.match(/ACP\s+REG_SZ\s+(\d+)/);
		if (match) {
			return ACP_TO_CODEPAGE[match[1]] ?? Codepage.Windows1252;
		}
	} catch {
		// Registry not readable (non-Windows or insufficient permissions)
	}

	return Codepage.Windows1252;
}

/**
 * Decode a buffer using the detected encoding.
 * BOM bytes are stripped from the output.
 */
export function decodeBuffer(buffer: Buffer, result?: SniffResult): string {
	const { encoding, hasBom } = result ?? sniffEncoding(buffer);

	if (encoding === "utf16le") {
		const bomOffset = hasBom ? 2 : 0;
		return buffer.subarray(bomOffset).toString("utf16le");
	}

	if (encoding === "utf16be") {
		const bomOffset = hasBom ? 2 : 0;
		return swap16Bytes(buffer.subarray(bomOffset)).toString("utf16le");
	}

	if (encoding === "utf8") {
		const bomOffset = hasBom ? 3 : 0;
		return buffer.subarray(bomOffset).toString("utf8");
	}

	// Fallback: use the system ANSI codepage via iconv-lite
	const label = codepageToLabel(getSystemCodepage());
	;
	return iconv.decode(buffer, label);
}

/**
 * Encode a JS string into a Buffer in the given codepage via
 * {@link https://www.npmjs.com/package/iconv-lite | iconv-lite}.
 * Throws if the codepage has no label mapping.
 */
export function encodeForCodepage(text: string, codepage: Codepage): Buffer {
	const label = codepageToLabel(codepage);
	if (!label) {
		throw new Error(`Cannot encode: no label for codepage ${codepage} (${Codepage[codepage]})`);
	}
	;
	return iconv.encode(text, label);
}

// ── helpers ──────────────────────────────────────────────────────

function swap16Bytes(buffer: Buffer): Buffer {
	const swapped = Buffer.allocUnsafe(buffer.length);
	for (let i = 0; i < buffer.length - 1; i += 2) {
		swapped[i] = buffer[i + 1];
		swapped[i + 1] = buffer[i];
	}
	if (buffer.length % 2 === 1) {
		swapped[buffer.length - 1] = buffer[buffer.length - 1];
	}
	return swapped;
}

/**
 * Strict UTF-8 validation.
 * Rejects overlong sequences, surrogate halves, and out-of-range codepoints.
 */
function isValidUtf8(buffer: Buffer): boolean {
	let i = 0;
	while (i < buffer.length) {
		const b0 = buffer[i];

		if (b0 < 0x80) {
			// ASCII
			i++;
			continue;
		}

		if ((b0 & 0xe0) === 0xc0) {
			// 2-byte sequence (U+0080 – U+07FF)
			if (i + 1 >= buffer.length) return false;
			const b1 = buffer[i + 1];
			if ((b1 & 0xc0) !== 0x80) return false;
			// Overlong: C0 and C1 are never valid
			if ((b0 & 0xfe) === 0xc0) return false;
			i += 2;
		} else if ((b0 & 0xf0) === 0xe0) {
			// 3-byte sequence (U+0800 – U+FFFF)
			if (i + 2 >= buffer.length) return false;
			const b1 = buffer[i + 1];
			const b2 = buffer[i + 2];
			if ((b1 & 0xc0) !== 0x80 || (b2 & 0xc0) !== 0x80) return false;
			// Overlong: E0 followed by 80–9F
			if (b0 === 0xe0 && (b1 & 0xe0) === 0x80) return false;
			// Surrogate halves: ED A0–BF
			if (b0 === 0xed && (b1 & 0xe0) === 0xa0) return false;
			i += 3;
		} else if ((b0 & 0xf8) === 0xf0) {
			// 4-byte sequence (U+10000 – U+10FFFF)
			if (i + 3 >= buffer.length) return false;
			const b1 = buffer[i + 1];
			const b2 = buffer[i + 2];
			const b3 = buffer[i + 3];
			if ((b1 & 0xc0) !== 0x80 || (b2 & 0xc0) !== 0x80 || (b3 & 0xc0) !== 0x80) return false;
			// Overlong: F0 followed by 80–8F
			if (b0 === 0xf0 && (b1 & 0xf0) === 0x80) return false;
			// Out of range: > U+10FFFF (F4 90+)
			if (b0 === 0xf4 && b1 > 0x8f) return false;
			if (b0 > 0xf4) return false;
			i += 4;
		} else {
			// Invalid start byte (0x80–0xBF, 0xF5–0xFF)
			return false;
		}
	}
	return true;
}
