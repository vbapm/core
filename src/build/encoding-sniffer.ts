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
	[Codepage.Windows1250]: "windows-1250",
	[Codepage.Windows1251]: "windows-1251",
	[Codepage.Windows1252]: "windows-1252",
	[Codepage.Windows1253]: "windows-1253",
	[Codepage.Windows1254]: "windows-1254",
	[Codepage.Windows1255]: "windows-1255",
	[Codepage.Windows1256]: "windows-1256",
	[Codepage.Windows1257]: "windows-1257",
	[Codepage.Windows1258]: "windows-1258"
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

/**
 * Return the local system ANSI codepage.
 *
 * When vbapm has just exported files from VBA (via `Component.Export`),
 * those files are guaranteed to be in this codepage. Pass this to the
 * Component constructor to skip the sniffing step entirely.
 */
export function getSystemCodepage(): Codepage {
	// VBA's `Component.Export` always writes in the system ANSI codepage.
	// On Western Windows machines this is windows-1252.
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

	// windows-1252 (no BOM possible for legacy codepages)
	return new TextDecoder("windows-1252").decode(buffer);
}

/**
 * Encode a JS string into a Buffer in the given codepage, for
 * writing .bas / .cls / .frm files that VBA's `Component.Import`
 * will read (VBA reads in the system ANSI codepage).
 *
 * Only {@link Codepage.Windows1252} is currently implemented.
 * {@link Codepage.Unknown} and {@link Codepage.UTF8} fall back to
 * writing UTF-8 with BOM (VBA can read UTF-8 with BOM).
 */
export function encodeForCodepage(text: string, codepage: Codepage): Buffer {
	if (codepage === Codepage.Windows1252) {
		return encodeAsWindows1252(text);
	}

	// UTF-8 with BOM — VBA can read UTF-8 files that have a BOM
	return Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(text, "utf8")]);
}

/** CP1252 characters in the 0x80-0x9F range that differ from latin1. */
const CP1252_80_9F: Record<number, number> = {
	0x20ac: 0x80, 0x201a: 0x82, 0x0192: 0x83, 0x201e: 0x84,
	0x2026: 0x85, 0x2020: 0x86, 0x2021: 0x87, 0x02c6: 0x88,
	0x2030: 0x89, 0x0160: 0x8a, 0x2039: 0x8b, 0x0152: 0x8c,
	0x017d: 0x8e, 0x2018: 0x91, 0x2019: 0x92, 0x201c: 0x93,
	0x201d: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97,
	0x02dc: 0x98, 0x2122: 0x99, 0x0161: 0x9a, 0x203a: 0x9b,
	0x0153: 0x9c, 0x017e: 0x9e, 0x0178: 0x9f,
};

function encodeAsWindows1252(text: string): Buffer {
	const buf = Buffer.alloc(text.length * 2);
	let out = 0;

	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i);
		const cp = CP1252_80_9F[code];
		if (cp !== undefined) {
			buf[out++] = cp;
		} else if (code < 0x100) {
			buf[out++] = code;
		} else {
			buf[out++] = 0x3f; // unmappable → '?'
		}
	}

	return buf.subarray(0, out);
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
