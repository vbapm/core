import { execFile } from "node:child_process";
import { mkdtemp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { promisify } from "node:util";
import { zip } from "../../utils/zip";

const execFileAsync = promisify(execFile);
const checker = resolve(process.cwd(), "scripts/check-power-query-mashup.mjs");
const formula = 'let Source = "match" in Source';

async function createFixture(sourceFormula: string) {
	const root = await mkdtemp(join(tmpdir(), "vbapm-mashup-check-test-"));
	const targetDirectory = join(root, "target");
	const sourceDirectory = join(root, "src", "pq");
	const packageDirectory = join(root, "package");
	const formulaDirectory = join(packageDirectory, "Formulas");
	const customXmlDirectory = join(targetDirectory, "customXml");

	await mkdir(formulaDirectory, { recursive: true });
	await mkdir(customXmlDirectory, { recursive: true });
	await mkdir(sourceDirectory, { recursive: true });
	await writeFile(
		join(formulaDirectory, "Section1.m"),
		`section Section1;\n\nshared #"Query" = ${formula};\n`,
		"utf8"
	);
	await writeFile(join(sourceDirectory, "Query.pq"), sourceFormula, "utf8");

	const archivePath = join(root, "mashup.zip");
	await zip(packageDirectory, archivePath);
	const packageBytes = await readFile(archivePath);
	const header = Buffer.alloc(8);
	header.writeUInt32LE(packageBytes.length, 4);
	const base64 = Buffer.concat([header, packageBytes]).toString("base64");
	await writeFile(
		join(customXmlDirectory, "item1.xml"),
		`<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">${base64}</DataMashup>`,
		"utf8"
	);

	return { root, targetDirectory, sourceDirectory };
}

async function runChecker(targetDirectory: string, sourceDirectory: string) {
	try {
		const result = await execFileAsync(process.execPath, [
			checker,
			targetDirectory,
			sourceDirectory
		]);
		return { code: 0, stdout: result.stdout, stderr: result.stderr };
	} catch (error) {
		const result = error as Error & { code: number; stdout: string; stderr: string };
		return { code: result.code, stdout: result.stdout, stderr: result.stderr };
	}
}

test("passes when the .pq formula matches the embedded Mashup formula", async () => {
	const fixture = await createFixture(formula);
	try {
		const result = await runChecker(fixture.targetDirectory, fixture.sourceDirectory);
		expect(result.code).toBe(0);
		expect(result.stdout).toContain("1 formulas match");
	} finally {
		await rm(fixture.root, { recursive: true, force: true });
	}
});

test("fails when the .pq formula differs from the embedded Mashup formula", async () => {
	const fixture = await createFixture('let Source = "different" in Source');
	try {
		const result = await runChecker(fixture.targetDirectory, fixture.sourceDirectory);
		expect(result.code).toBe(1);
		expect(result.stderr).toContain('No .pq formula matches embedded query "Query".');
		expect(result.stderr).toContain("No embedded query matches source");
	} finally {
		await rm(fixture.root, { recursive: true, force: true });
	}
});
