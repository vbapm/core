import { Buffer } from "buffer";
import {
	Codepage,
	codepageToLabel,
	decodeBuffer,
	getSystemCodepage,
	sniffEncoding
} from "../encoding-sniffer";

// ── helpers ──────────────────────────────────────────────────────

/** Build a CP1252-encoded buffer from a JS string with French accents. */
function cp1252(text: string): Buffer {
	const MAP: Record<string, number> = {
		"€": 0x80,
		"‚": 0x82,
		ƒ: 0x83,
		"„": 0x84,
		"…": 0x85,
		"†": 0x86,
		"‡": 0x87,
		ˆ: 0x88,
		"‰": 0x89,
		Š: 0x8a,
		"‹": 0x8b,
		Œ: 0x8c,
		Ž: 0x8e,
		"‘": 0x91,
		"’": 0x92,
		"“": 0x93,
		"”": 0x94,
		"•": 0x95,
		"–": 0x96,
		"—": 0x97,
		"˜": 0x98,
		"™": 0x99,
		š: 0x9a,
		"›": 0x9b,
		œ: 0x9c,
		ž: 0x9e,
		Ÿ: 0x9f,
		é: 0xe9,
		è: 0xe8,
		ê: 0xea,
		ë: 0xeb,
		à: 0xe0,
		â: 0xe2,
		ä: 0xe4,
		ù: 0xf9,
		û: 0xfb,
		ü: 0xfc,
		ç: 0xe7,
		É: 0xc9,
		È: 0xc8,
		Ê: 0xca,
		Ë: 0xcb,
		À: 0xc0,
		Â: 0xc2,
		Ä: 0xc4,
		Ù: 0xd9,
		Û: 0xdb,
		Ü: 0xdc,
		Ç: 0xc7,
		ô: 0xf4,
		î: 0xee,
		ï: 0xef,
		ñ: 0xf1
	};

	const bytes: number[] = [];
	for (let i = 0; i < text.length; i++) {
		const ch = text[i];
		const cp = MAP[ch];
		bytes.push(cp !== undefined ? cp : ch.charCodeAt(0));
	}
	return Buffer.from(bytes);
}

// ── sniffEncoding ────────────────────────────────────────────────

describe("sniffEncoding", () => {
	test("empty buffer → utf8 without BOM", () => {
		expect(sniffEncoding(Buffer.alloc(0))).toEqual({
			encoding: "utf8",
			hasBom: false
		});
	});

	test("ASCII-only → utf8 without BOM", () => {
		const buf = Buffer.from("Hello World", "ascii");
		expect(sniffEncoding(buf)).toEqual({ encoding: "utf8", hasBom: false });
	});

	test("valid UTF-8 (French via proper encoding) → utf8 without BOM", () => {
		const buf = Buffer.from("é è ê à ç ù", "utf8");
		expect(sniffEncoding(buf)).toEqual({ encoding: "utf8", hasBom: false });
	});

	test("UTF-8 BOM → utf8 with BOM", () => {
		const buf = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from("test", "utf8")]);
		expect(sniffEncoding(buf)).toEqual({ encoding: "utf8", hasBom: true });
	});

	test("UTF-16 LE BOM → utf16le with BOM", () => {
		const buf = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from("test", "utf16le")]);
		expect(sniffEncoding(buf)).toEqual({ encoding: "utf16le", hasBom: true });
	});

	test("UTF-16 BE BOM → utf16be with BOM", () => {
		const buf = Buffer.concat([Buffer.from([0xfe, 0xff]), Buffer.from("test", "utf16le").swap16()]);
		expect(sniffEncoding(buf)).toEqual({ encoding: "utf16be", hasBom: true });
	});

	test("CP1252 with French accents → windows-1252 (not valid UTF-8)", () => {
		const buf = cp1252("é è ê à ç ù");
		expect(sniffEncoding(buf)).toEqual({
			encoding: "windows-1252",
			hasBom: false
		});
	});

	test("CP1252 with euro sign (0x80) → windows-1252", () => {
		const buf = Buffer.from([0x80]); // € in CP1252, invalid in UTF-8
		expect(sniffEncoding(buf)).toEqual({
			encoding: "windows-1252",
			hasBom: false
		});
	});

	test("buffer starting with 0xC0 (overlong, invalid UTF-8) → windows-1252", () => {
		const buf = Buffer.from("Hello\xC0World", "latin1");
		expect(sniffEncoding(buf)).toEqual({
			encoding: "windows-1252",
			hasBom: false
		});
	});

	test("buffer with truncated 2-byte UTF-8 sequence → windows-1252", () => {
		const buf = Buffer.from("Hello\xC2", "latin1"); // 0xC2 needs a continuation byte
		expect(sniffEncoding(buf)).toEqual({
			encoding: "windows-1252",
			hasBom: false
		});
	});

	test("UTF-16 LE without BOM (sufficient null bytes) → utf16le", () => {
		// "A" in UTF-16LE = 0x41 0x00, repeat enough times
		const chars = "A".repeat(20);
		const buf = Buffer.from(chars, "utf16le");
		expect(sniffEncoding(buf)).toEqual({ encoding: "utf16le", hasBom: false });
	});
});

// ── decodeBuffer ─────────────────────────────────────────────────

describe("decodeBuffer", () => {
	test("ASCII-only → correct text", () => {
		const buf = Buffer.from("Hello World", "ascii");
		expect(decodeBuffer(buf)).toBe("Hello World");
	});

	test("UTF-8 with BOM → BOM stripped", () => {
		const buf = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from("Bonjour", "utf8")]);
		expect(decodeBuffer(buf)).toBe("Bonjour");
	});

	test("UTF-8 without BOM → correct text", () => {
		const buf = Buffer.from("Café crème", "utf8");
		expect(decodeBuffer(buf)).toBe("Café crème");
	});

	test("UTF-16 LE with BOM → BOM stripped, correct text", () => {
		const buf = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from("Hello", "utf16le")]);
		expect(decodeBuffer(buf)).toBe("Hello");
	});

	test("CP1252 French → correct accents (no replacement chars)", () => {
		const buf = cp1252("é è ê à ç ù É È Ê À Ç Ù");
		const result = decodeBuffer(buf);
		expect(result).toBe("é è ê à ç ù É È Ê À Ç Ù");
		expect(result).not.toContain("\uFFFD");
	});

	test("CP1252 VBA module-like content → correct", () => {
		const buf = cp1252('\' Module de démonstration\nMsgBox "Voilà !"');
		const result = decodeBuffer(buf);
		expect(result).toBe('\' Module de démonstration\nMsgBox "Voilà !"');
		expect(result).not.toContain("\uFFFD");
	});

	test("CP1252 œ (0x9C) and Œ (0x8C) → correct", () => {
		const buf = Buffer.from([0x9c, 0x8c]); // œ Œ
		const result = decodeBuffer(buf);
		expect(result).toBe("œŒ");
		expect(result).not.toContain("\uFFFD");
	});

	test("CP1252 Euro sign (0x80) and bullet (0x95) → correct", () => {
		const buf = Buffer.from([0x80, 0x95]); // € •
		const result = decodeBuffer(buf);
		expect(result).toBe("€•");
		expect(result).not.toContain("\uFFFD");
	});

	test("pre-computed SniffResult skips sniffing", () => {
		const buf = cp1252("éàç");
		const result = decodeBuffer(buf, { encoding: "windows-1252", hasBom: false });
		expect(result).toBe("éàç");
	});

	test("UTF-16 BE with BOM → correct text", () => {
		const content = "Hello";
		const utf16le = Buffer.from(content, "utf16le");
		const utf16be = Buffer.alloc(utf16le.length);
		for (let i = 0; i < utf16le.length - 1; i += 2) {
			utf16be[i] = utf16le[i + 1];
			utf16be[i + 1] = utf16le[i];
		}
		const buf = Buffer.concat([Buffer.from([0xfe, 0xff]), utf16be]);
		expect(decodeBuffer(buf)).toBe("Hello");
	});
});

// ── codepageToLabel ──────────────────────────────────────────────

describe("codepageToLabel", () => {
	test("Unknown → empty string", () => {
		expect(codepageToLabel(Codepage.Unknown)).toBe("");
	});

	for (const cp of [
		Codepage.Windows1250,
		Codepage.Windows1251,
		Codepage.Windows1252,
		Codepage.Windows1253,
		Codepage.Windows1254,
		Codepage.Windows1255,
		Codepage.Windows1256,
		Codepage.Windows1257,
		Codepage.Windows1258
	]) {
		test(`${Codepage[cp]} → "windows-${cp}"`, () => {
			expect(codepageToLabel(cp)).toBe(`windows-${cp}`);
		});
	}

	test("UTF8 → utf-8", () => {
		expect(codepageToLabel(Codepage.UTF8)).toBe("utf-8");
	});

	test("Windows932 → cp932", () => {
		expect(codepageToLabel(Codepage.Windows932)).toBe("cp932");
	});

	test("Windows936 → gbk", () => {
		expect(codepageToLabel(Codepage.Windows936)).toBe("gbk");
	});

	test("Windows874 → windows-874", () => {
		expect(codepageToLabel(Codepage.Windows874)).toBe("windows-874");
	});

	test("Windows949 → ks_c_5601-1987", () => {
		expect(codepageToLabel(Codepage.Windows949)).toBe("ks_c_5601-1987");
	});

	test("Windows950 → big5", () => {
		expect(codepageToLabel(Codepage.Windows950)).toBe("big5");
	});
});

// ── getSystemCodepage ────────────────────────────────────────────

describe("getSystemCodepage", () => {
	test("returns a known Codepage (reads ACP from registry)", () => {
		const cp = getSystemCodepage();
		// Should return a known value; Windows1252 is the default fallback.
		expect(Object.values(Codepage)).toContain(cp);
		expect(cp).not.toBe(Codepage.Unknown);
	});
});

// ── Roundtrip: sniff → decode (end-to-end) ───────────────────────

describe("sniff → decode roundtrip", () => {
	test("UTF-8 roundtrip", () => {
		const original = "Café crème à la française";
		const buf = Buffer.from(original, "utf8");
		const result = decodeBuffer(buf);
		expect(result).toBe(original);
	});

	test("CP1252 roundtrip via sniff", () => {
		const original = "é è ê à ç ù œ";
		const buf = cp1252(original);
		const result = decodeBuffer(buf);
		expect(result).toBe(original);
	});

	test("no replacement characters in CP1252 VBA source", () => {
		const vba = [
			'Attribute VB_Name = "ModuleAccents"',
			"",
			"' Démonstration",
			"Public Sub Test()",
			'    MsgBox "Voilà !"',
			"End Sub"
		].join("\n");
		const buf = cp1252(vba);
		const result = decodeBuffer(buf);
		expect(result).toBe(vba);
		expect(result).not.toContain("\uFFFD");
	});
});
