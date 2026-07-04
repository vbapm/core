import { detectImportEncoding } from "../detect-encoding";
import { join } from "../../utils/path";
import { dir } from "../../../tests/__fixtures__";
import { writeFileSync, mkdirSync, rmSync } from "fs";

const fixtureDir = join(dir, "projects", "detect-encoding-test");

beforeAll(() => {
	mkdirSync(fixtureDir, { recursive: true });
});

afterAll(() => {
	rmSync(fixtureDir, { recursive: true, force: true });
});

describe("detectImportEncoding", () => {
	// Valid Windows codepage labels that detectImportEncoding may return
	const validLabels = new Set([
		"windows-1250", "windows-1251", "windows-1252", "windows-1253",
		"windows-1254", "windows-1255", "windows-1256", "windows-1257", "windows-1258",
		"windows-932", "windows-936", "windows-874", "windows-949", "windows-950",
		"cp1250", "cp1251", "cp1252",
		"cp932", "cp936", "cp949", "cp950", "cp874"
	]);

	test("returns undefined for ASCII-only file", async () => {
		const path = join(fixtureDir, "ascii.bas");
		writeFileSync(path, 'Attribute VB_Name = "Module1"\nPublic Sub Hello()\nEnd Sub\n');

		const result = await detectImportEncoding(path);
		expect(result).toBeUndefined();
	});

	test("detects system codepage for non-ASCII file", async () => {
		const path = join(fixtureDir, "cp1252.bas");
		const iconv = require("iconv-lite");
		const buf = iconv.encode(
			'Attribute VB_Name = "Module1"\n\' Déjà vu\nPublic Sub Hello()\nEnd Sub\n',
			"windows-1252"
		);
		writeFileSync(path, buf);

		const result = await detectImportEncoding(path);
		expect(result).toBeDefined();
		expect(validLabels.has(result!)).toBe(true);
	});

	test("detects encoding for Windows-932 content", async () => {
		const path = join(fixtureDir, "cp932.bas");
		const iconv = require("iconv-lite");
		// Longer Japanese sample — jschardet needs enough non-ASCII bytes
		// to confidently distinguish Windows-932 from single-byte codepages.
		const buf = iconv.encode(
			'Attribute VB_Name = "Module1"\n' +
			"' こんにちは世界\n" +
			"' これは日本語のテストです\n" +
			"' 漢字・ひらがな・カタカナ\n" +
			"' 文字化けを防ぐためにエンコーディングを指定します\n" +
			"Public Sub Hello()\n" +
			'  MsgBox "こんにちは"\n' +
			"End Sub\n",
			"windows-932"
		);
		writeFileSync(path, buf);

		const result = await detectImportEncoding(path);
		expect(result).toBeDefined();
		expect(validLabels.has(result!)).toBe(true);
		expect(result!.toLowerCase()).toMatch(/^(windows-?932|cp932)$/);
		// jschardet is ESM-only; in Jest's CJS environment the dynamic
		// import fails silently and we fall back to the system codepage.
		// On a Japanese-locale machine (or in real CLI usage) this would
		// return "cp932". The validLabels check above covers both cases.
	});
});
