import { writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import * as iconv from "iconv-lite";

const CODEPAGES = [1252, 1251, 1250, 932, 936, 874, 949, 950, 1253, 1254, 1255, 1256, 1257, 1258];
const EXPECTED_COMMON_CODEPOINTS = [0x2018, 0x2019, 0x201c, 0x201d, 0x2026];
const cp1252Label = "windows-1252";

function roundTrips(value, codepage) {
	const label = `windows-${codepage}`;
	return iconv.decode(iconv.encode(value, label), label) === value;
}

function formatCodepoint(codepoint) {
	return `U+${codepoint.toString(16).toUpperCase().padStart(4, "0")}`;
}

function getEncodedBytes(codepoint, codepage) {
	const value = String.fromCodePoint(codepoint);
	const label = `windows-${codepage}`;
	const bytes = iconv.encode(value, label);
	if (iconv.decode(bytes, label) !== value) return "not representable";
	return [...bytes].map(byte => byte.toString(16).toUpperCase().padStart(2, "0")).join(" ");
}

function createMarkdownReport() {
	const headers = ["Character", "Code point", ...CODEPAGES.map(codepage => `CP${codepage}`)];
	const commonRows = commonCodepoints.map(codepoint =>
		[
			cp1252Characters.get(codepoint),
			formatCodepoint(codepoint),
			...CODEPAGES.map(codepage => getEncodedBytes(codepoint, codepage))
		].join(" | ")
	);
	const emDashRow = [
		"—",
		formatCodepoint(0x2014),
		...CODEPAGES.map(codepage => getEncodedBytes(0x2014, codepage))
	].join(" | ");
	const bulletRow = [
		"•",
		formatCodepoint(0x2022),
		...CODEPAGES.map(codepage => getEncodedBytes(0x2022, codepage))
	].join(" | ");

	return [
		"# Characters Shared by the Multilingual Test Codepages",
		"",
		`The scan found **${commonCodepoints.length} non-ASCII characters** that round-trip through CP1252 and all ${CODEPAGES.length - 1} other codepages in this analysis.`,
		"",
		`Codepages checked, in order: ${CODEPAGES.map(codepage => `CP${codepage}`).join(", ")}.`,
		"The set includes the eight codepages with multilingual fixtures plus Windows-1253 through Windows-1258. The Excel build/export loop still runs only for codepages with fixtures.",
		"ASCII characters are not listed; the report focuses on non-ASCII characters.",
		"",
		"## Method",
		"",
		`The [Node.js scanner](../../../../scripts/check-common-windows-codepage-characters.mjs) enumerates all 128 high-byte values in CP1252. It retains each decoded character only if encoding and decoding it returns the exact same character, then checks that set against every codepage above using the same round-trip test. This checks all **${cp1252Characters.size}** non-ASCII CP1252 characters that round-trip.`,
		"",
		"The [PowerShell verifier](../../../../scripts/ps/Test-CommonWindowsCodepageCharacters.ps1) independently checks the listed characters with strict .NET Windows codepage encoders and decoders, and prints their byte values.",
		"",
		"## Common Non-ASCII Characters",
		"",
		`| ${headers.join(" | ")} |`,
		`| ${headers.map(() => "---").join(" | ")} |`,
		...commonRows.map(row => `| ${row} |`),
		"",
		"## Em Dash",
		"",
		`U+2014 (em dash) is not common to all ${CODEPAGES.length} codepages. It cannot be represented in CP932 or CP949.`,
		"",
		`| ${headers.join(" | ")} |`,
		`| ${headers.map(() => "---").join(" | ")} |`,
		`| ${emDashRow} |`,
		"",
		"## Bullet",
		"",
		`U+2022 (bullet) is not common to all ${CODEPAGES.length} codepages. It cannot be represented in CP932, CP936, CP949, or CP950.`,
		"",
		`| ${headers.join(" | ")} |`,
		`| ${headers.map(() => "---").join(" | ")} |`,
		`| ${bulletRow} |`,
		"",
		"> This report establishes codepage round-trip support only. It does not show that VBA accepts these punctuation characters in identifiers.",
		"",
		"## Reproduce",
		"",
		"From the repository root, run:",
		"",
		"```powershell",
		"node scripts/check-common-windows-codepage-characters.mjs --markdown docs/analysis/codepages/commonality/windows-codepage-character-report.md",
		".\\scripts\\ps\\Test-CommonWindowsCodepageCharacters.ps1",
		"```",
		"",
		`The [multilingual GitHub Actions workflow](../../../../.github/workflows/multilingual-encoding-test.yml) continues to run the Excel build/export fixtures; these wider codepage scans are manual checks.`,
		""
	].join("\n");
}

const cp1252Characters = new Map();
for (let byte = 0x80; byte <= 0xff; byte++) {
	const value = iconv.decode(Buffer.from([byte]), cp1252Label);
	if ([...value].length !== 1) continue;

	const codepoint = value.codePointAt(0);
	if (codepoint > 0x7f && roundTrips(value, 1252)) {
		cp1252Characters.set(codepoint, value);
	}
}

const commonCodepoints = [...cp1252Characters.keys()]
	.filter(codepoint => {
		const value = cp1252Characters.get(codepoint);
		return CODEPAGES.every(codepage => roundTrips(value, codepage));
	})
	.sort((left, right) => left - right);

console.log(`Codepages checked: ${CODEPAGES.join(", ")}`);
console.log(
	`Method: enumerate every non-ASCII CP1252 byte that round-trips, then test each character against all ${CODEPAGES.length} codepages.`
);
console.log(`Non-ASCII CP1252 round-trip characters checked: ${cp1252Characters.size}`);
console.log(`Common non-ASCII characters: ${commonCodepoints.length}`);
console.log(["Code point", "Character", ...CODEPAGES.map(codepage => `CP${codepage}`)].join("\t"));
for (const codepoint of commonCodepoints) {
	console.log(
		[
			formatCodepoint(codepoint),
			cp1252Characters.get(codepoint),
			...CODEPAGES.map(codepage => getEncodedBytes(codepoint, codepage))
		].join("\t")
	);
}

console.log("\nEm dash (U+2014) support:");
for (const codepage of CODEPAGES) {
	console.log(`CP${codepage}: ${getEncodedBytes(0x2014, codepage)}`);
}

console.log("\nBullet (U+2022) support:");
for (const codepage of CODEPAGES) {
	console.log(`CP${codepage}: ${getEncodedBytes(0x2022, codepage)}`);
}

if (commonCodepoints.join(",") !== EXPECTED_COMMON_CODEPOINTS.join(",")) {
	throw new Error(
		`Unexpected common-character set: ${commonCodepoints.map(formatCodepoint).join(", ")}`
	);
}

const args = process.argv.slice(2);
if (args.length > 0) {
	if (args.length !== 2 || args[0] !== "--markdown" || !args[1]) {
		throw new Error(
			"Usage: node scripts/check-common-windows-codepage-characters.mjs [--markdown <output.md>]"
		);
	}

	const outputPath = resolve(args[1]);
	await writeFile(outputPath, createMarkdownReport(), "utf8");
	console.log(`\nMarkdown report written to ${outputPath}`);
}
