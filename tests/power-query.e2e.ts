/**
 * End-to-end tests for Power Query source updates during build.
 */

import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { readFile, writeFile } from "fs-extra";
import { join, resolve } from "path";
import { powerQuery } from "./__fixtures__";
import { closePersistentSession, execute, setup } from "./__helpers__/execute";

const execFileAsync = promisify(execFile);
const checkMashupScript = resolve(process.cwd(), "scripts/check-power-query-mashup.mjs");

process.env.E2E_BACKGROUND = "1";
jest.setTimeout(180000);

afterAll(async () => {
	await closePersistentSession();
});

async function expectMashupMatchesSource(cwd: string, queryCount = 1) {
	const { stdout } = await execFileAsync(process.execPath, [
		checkMashupScript,
		join(cwd, "target"),
		join(cwd, "src")
	]);
	expect(stdout).toContain(`Power Query check passed: ${queryCount} formulas match.`);
}

describe("Power Query build", () => {
	test("builds when the source formula matches the template query", async () => {
		await setup(powerQuery, "power-query-match", async cwd => {
			await execute(cwd, "build");
			await execute(cwd, "extract --target xlsm");

			await expectMashupMatchesSource(cwd);
		});
	});

	test("updates an existing query from its .pq source", async () => {
		await setup(powerQuery, "power-query-update", async cwd => {
			const formula = [
				"let",
				'    Source = #table({"Column1", "Column2"}, {{"e2e-update", "ok"}})',
				"in",
				"    Source"
			].join("\n");
			await writeFile(join(cwd, "src", "pq", "Table 1.pq"), formula, "utf8");

			await execute(cwd, "build");
			await execute(cwd, "extract --target xlsm");

			const extractedFormula = await readFile(join(cwd, "src", "pq", "Table 1.pq"), "utf8");
			expect(extractedFormula).toContain('"e2e-update"');
			await expectMashupMatchesSource(cwd);
		});
	});

	test("creates a workbook query from a source added with vba add", async () => {
		await setup(powerQuery, "power-query-add", async cwd => {
			await execute(cwd, "add NewQuery.pq");
			const formula = [
				"let",
				'    Source = #table({"Value"}, {{"added-query"}})',
				"in",
				"    Source"
			].join("\n");
			await writeFile(join(cwd, "src", "NewQuery.pq"), formula, "utf8");

			await execute(cwd, "build");
			await execute(cwd, "extract --target xlsm");

			const extractedFormula = await readFile(join(cwd, "src", "NewQuery.pq"), "utf8");
			expect(extractedFormula).toContain('"added-query"');
			await expectMashupMatchesSource(cwd, 2);
		});
	});
});

test("preserves a reserved query name when extraction sanitizes its filename", async () => {
	await setup(powerQuery, "power-query-reserved-name", async cwd => {
		const manifestPath = join(cwd, "vbaproject.toml");
		const manifest = await readFile(manifestPath, "utf8");
		await writeFile(
			manifestPath,
			manifest.replace(
				'PowerQueries = "src/pq/*.pq"',
				'NUL = "src/pq/_NUL.pq"\nPowerQueries = "src/pq/*.pq"'
			),
			"utf8"
		);
		await writeFile(join(cwd, "src", "pq", "_NUL.pq"), "let Source = 42 in Source", "utf8");

		await execute(cwd, "build");
		await execute(cwd, "extract --target xlsm");

		expect(await readFile(join(cwd, "src", "pq", "_NUL.pq"), "utf8")).toContain("Source = 42");
		const extractedManifest = await readFile(manifestPath, "utf8");
		expect(extractedManifest).toContain('NUL = "src/pq/_NUL.pq"');
		await expectMashupMatchesSource(cwd, 2);
	});
});
