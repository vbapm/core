/**
 * Multilingual encoding tests for the `vba` CLI.
 *
 * These tests verify that VBA source files encoded in different Windows
 * ANSI codepages survive the full roundtrip: build → export.
 *
 * ## Environment
 *
 * The test requires the `E2E_ML` environment variable to be set to `1`,
 * `true`, or `yes`. On a developer machine without the correct system
 * locale, the test is skipped.
 *
 * In CI, each codepage fixture runs on a Windows image configured with
 * the matching system locale (e.g., `fr-FR` for CP1252, `ru-RU` for
 * CP1251, `pl-PL` for CP1250).
 *
 * ## Fixture layout
 *
 * `tests/__fixtures__/multilingual/<codepage>/` — each directory is a
 * minimal vbapm project with one `.bas` module containing characters
 * specific to that codepage.
 *
 * ## Codepage detection
 *
 * The current system ANSI codepage is read from the registry at
 * `HKLM\SYSTEM\CurrentControlSet\Control\Nls\CodePage\ACP`. Only the
 * fixture matching the current codepage is tested.
 *
 * ## Running
 *
 *   pnpm test:e2e:multilang
 */

import { readFile } from "fs-extra";
import { join } from "path";
import { execSync } from "node:child_process";
import iconv from "iconv-lite";
import { execute, setup } from "./__helpers__/execute";

const isMultilingualTest =
	/^(1|true|yes)$/i.test(process.env.E2E_ML || "") || /^(1|true|yes)$/i.test(process.env.CI || "");

// Only run on CI or when explicitly enabled
const describeML = isMultilingualTest ? describe : describe.skip;

/**
 * Read a file using iconv-lite with the system ANSI codepage.
 * After vba export, the src/ file is still in the system codepage
 * (applyChangeset skips writing unchanged components), so reading
 * as UTF-8 would produce replacement characters.
 */
async function readExportedFile(dir: string, filename: string): Promise<string> {
	const { getSystemCodepage, codepageToLabel } = await import("../src/build/encoding-sniffer");
	const buffer = await readFile(join(dir, "src", filename));
	return iconv.decode(buffer, codepageToLabel(getSystemCodepage()));
}

/**
 * Read the system ANSI codepage from the Windows registry.
 * Returns the codepage number as a string (e.g. "1252"), or undefined
 * if it cannot be read.
 */
function getSystemAnsiCodepage(): string | undefined {
	try {
		const result = execSync(
			"powershell -Command \"(Get-ItemProperty 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Nls\\CodePage').ACP\"",
			{ encoding: "utf8", timeout: 5000 }
		);
		return result.trim();
	} catch {
		try {
			const result = execSync(
				'reg query "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Nls\\CodePage" /v ACP',
				{ encoding: "utf8", timeout: 5000 }
			);
			const match = result.match(/ACP\s+REG_SZ\s+(\d+)/);
			return match ? match[1] : undefined;
		} catch {
			return undefined;
		}
	}
}

/**
 * Verify a VBA source file contains the expected characters (no �).
 */
async function verifyAccents(
	dir: string,
	filename: string,
	expectedStrings: string[]
): Promise<void> {
	const content = await readExportedFile(dir, filename);

	// No replacement characters
	expect(content).not.toContain("\uFFFD");

	for (const str of expectedStrings) {
		expect(content).toContain(str);
	}
}

jest.setTimeout(180000);

// ── Fixtures ────────────────────────────────────────────────────

interface MultilingualFixture {
	/** Codepage number (e.g. 1252) */
	codepage: string;
	/** Path to the fixture directory */
	dir: string;
	/** Name of the .bas source file to verify */
	filename: string;
	/** Strings that must appear in the exported source */
	expectedStrings: string[];
}

const fixturesDir = join(process.cwd(), "tests", "__fixtures__", "multilingual");

const FIXTURES: MultilingualFixture[] = [
	{
		codepage: "1252",
		dir: join(fixturesDir, "cp1252"),
		filename: "French.bas",
		expectedStrings: [
			"é è ê ë à â ä ù û ü ç",
			"É È Ê Ë À Â Ä Ù Û Ü Ç",
			"ä ö ü ß Ä Ö Ü",
			"á é í ó ú ñ",
			"Voilà les caractères accentués",
			"caractères"
		]
	},
	{
		codepage: "1251",
		dir: join(fixturesDir, "cp1251"),
		filename: "Russian.bas",
		expectedStrings: ["Привет мир", "А Б В Г Д Е"]
	},
	{
		codepage: "1250",
		dir: join(fixturesDir, "cp1250"),
		filename: "Polish.bas",
		expectedStrings: ["ą ć ę ł ń ó ś ź ż", "Ą Ć Ę Ł Ń Ó Ś Ź Ż", "Witaj świecie"]
	},
	{
		codepage: "1253",
		dir: join(fixturesDir, "cp1253"),
		filename: "Greek.bas",
		expectedStrings: ["Ελληνικά", "Καλημέρα κόσμε!"]
	},
	{
		codepage: "1254",
		dir: join(fixturesDir, "cp1254"),
		filename: "Turkish.bas",
		expectedStrings: ["Türkçe: Ç Ğ İ Ö Ş Ü ç ğ ı ö ş ü", "Merhaba dünya!"]
	},
	{
		codepage: "1255",
		dir: join(fixturesDir, "cp1255"),
		filename: "Hebrew.bas",
		expectedStrings: ["עברית", "שלום עולם!"]
	},
	{
		codepage: "1256",
		dir: join(fixturesDir, "cp1256"),
		filename: "Arabic.bas",
		expectedStrings: ["اللغة العربية", "مرحبا بالعالم!"]
	},
	{
		codepage: "1257",
		dir: join(fixturesDir, "cp1257"),
		filename: "Lithuanian.bas",
		expectedStrings: [
			"Latviešu: ā č ē ģ ī ķ ļ ņ š ū ž",
			"Lietuvių: ą č ę ė į š ų ū ž",
			"Sveiki, lietuvių pasauli!"
		]
	},
	{
		codepage: "1258",
		dir: join(fixturesDir, "cp1258"),
		filename: "Vietnamese.bas",
		expectedStrings: ["Vietnamese: Đ đ Ă ă Ơ ơ Ư ư ₫", "Xin chào!"]
	},
	{
		codepage: "932",
		dir: join(fixturesDir, "cp932"),
		filename: "Japanese.bas",
		expectedStrings: ["こんにちは世界", "日本語", "漢字"]
	},
	{
		codepage: "936",
		dir: join(fixturesDir, "cp936"),
		filename: "SimplifiedChinese.bas",
		expectedStrings: ["你好世界", "简体中文", "汉字"]
	},
	{
		codepage: "874",
		dir: join(fixturesDir, "cp874"),
		filename: "Thai.bas",
		expectedStrings: ["สวัสดี", "ไทย", "ภาษาไทย"]
	},
	{
		codepage: "949",
		dir: join(fixturesDir, "cp949"),
		filename: "Korean.bas",
		expectedStrings: ["안녕하세요", "한국어", "한글"]
	},
	{
		codepage: "950",
		dir: join(fixturesDir, "cp950"),
		filename: "TraditionalChinese.bas",
		expectedStrings: ["你好世界", "繁體中文", "漢字", "臺灣"]
	}
];

// ── Tests ───────────────────────────────────────────────────────

describeML("multilingual encoding", () => {
	const currentCodepage = getSystemAnsiCodepage();

	if (!currentCodepage) {
		test.skip("cannot read system codepage", () => {});
		return;
	}

	const matchingFixture = FIXTURES.find(f => f.codepage === currentCodepage);

	if (!matchingFixture) {
		test.skip(
			`no fixture for system codepage ${currentCodepage} ` +
				`(available: ${FIXTURES.map(f => f.codepage).join(", ")})`,
			() => {}
		);
		return;
	}

	test(
		`roundtrip preserves ${matchingFixture.codepage} characters ` +
			`(${matchingFixture.dir.split("/").pop()})`,
		async () => {
			await setup(matchingFixture.dir, `ml-${matchingFixture.codepage}`, async cwd => {
				// Ensure Excel is closed before changing codepage
				await closeExcel();

				// 1. Build the project (imports .bas into .xlsm)
				await execute(cwd, "build");

				// 2. Export to extract VBA source back
				await execute(cwd, "export --target xlsm");

				// 3. Verify the exported source has correct characters
				await verifyAccents(cwd, matchingFixture.filename, matchingFixture.expectedStrings);
			});
		}
	);

	// This fixture intentionally lacks encoding in [source] and should fail the build
	test("non-ASCII without encoding fails with helpful error", async () => {
		const noEncDir = join(fixturesDir, "no-encoding");

		await setup(noEncDir, "ml-no-enc", async cwd => {
			await expect(execute(cwd, "build")).rejects.toThrow(/Non-ASCII characters detected/);
		});
	});
});

/**
 * Close any running Excel instances to ensure a clean state for
 * codepage-sensitive operations.
 */
async function closeExcel(): Promise<void> {
	try {
		execSync("taskkill /f /im excel.exe 2>nul", { timeout: 5000 });
	} catch {
		// Excel may not be running — that's fine
	}
	// Wait for COM to release
	await new Promise(resolve => setTimeout(resolve, 2000));
}
