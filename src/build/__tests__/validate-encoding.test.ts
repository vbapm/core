import { loadFromProject } from "../load-from-project";
import { reset, setup } from "../../../tests/__helpers__/project";
import { dir } from "../../../tests/__fixtures__";
import { join } from "../../utils/path";
import { writeFileSync, mkdirSync, rmSync } from "fs";

afterEach(reset);

describe("encoding validation", () => {
	const fixtureDir = join(dir, "projects", "encoding-test");

	beforeAll(() => {
		mkdirSync(join(fixtureDir, "src"), { recursive: true });
		mkdirSync(join(fixtureDir, "targets"), { recursive: true });
	});

	afterAll(() => {
		rmSync(fixtureDir, { recursive: true, force: true });
	});

	test("ASCII-only source without encoding → builds fine", async () => {
		const toml = `\
[package]
name = "ascii-only"
version = "1.0.0"
authors = ["Test"]

[src]
Module1 = "src/Module1.bas"
`;

		const srcPath = join(fixtureDir, "src", "Module1.bas");
		const manifestPath = join(fixtureDir, "vbaproject.toml");

		await writeFileSync(srcPath, 'Attribute VB_Name = "Module1"\nPublic Sub Hello()\n  MsgBox "Hi"\nEnd Sub\n');
		await writeFileSync(manifestPath, toml);

		const { project, dependencies } = await setup(fixtureDir);
		const graph = await loadFromProject(project, dependencies);
		expect(graph.components.length).toBe(1);
		expect(graph.components[0].code).toContain("Hello");
	});

	test("non-ASCII without encoding → fails with suggestion", async () => {
		const toml = `\
[package]
name = "non-ascii-no-enc"
version = "1.0.0"
authors = ["Test"]

[src]
Module1 = "src/Module1.bas"
`;

		const srcPath = join(fixtureDir, "src", "Module1.bas");
		const manifestPath = join(fixtureDir, "vbaproject.toml");

		// French accented characters
		const srcContent = 'Attribute VB_Name = "Module1"\n\' Déjà vu – naïve façade\nPublic Sub Hello()\n  MsgBox "é"  \' café\nEnd Sub\n';
		writeFileSync(srcPath, srcContent);
		writeFileSync(manifestPath, toml);

		const { project, dependencies } = await setup(fixtureDir);
		await expect(loadFromProject(project, dependencies))
			.rejects.toThrow(/Non-ASCII characters detected/);
	});

	test("non-ASCII with src-encoding → builds fine", async () => {
		const toml = `\
[package]
name = "non-ascii-enc"
version = "1.0.0"
authors = ["Test"]
src-encoding = "cp1252"

[src]
Module1 = "src/Module1.bas"
`;

		const srcPath = join(fixtureDir, "src", "Module1.bas");
		const manifestPath = join(fixtureDir, "vbaproject.toml");

		const srcContent = 'Attribute VB_Name = "Module1"\n\' Déjà vu – naïve façade\nPublic Sub Hello()\n  MsgBox "é"  \' café\nEnd Sub\n';
		writeFileSync(srcPath, srcContent);
		writeFileSync(manifestPath, toml);

		const { project, dependencies } = await setup(fixtureDir);
		const graph = await loadFromProject(project, dependencies);
		expect(graph.components.length).toBe(1);
	});

	test("per-source encoding overrides project-level", async () => {
		const toml = `\
[package]
name = "per-source-enc"
version = "1.0.0"
authors = ["Test"]
src-encoding = "cp1252"

[src]
Module1 = { path = "src/Module1.bas", encoding = "cp932" }
`;

		const srcPath = join(fixtureDir, "src", "Module1.bas");
		const manifestPath = join(fixtureDir, "vbaproject.toml");

		// Japanese characters (CP932)
		const srcContent = 'Attribute VB_Name = "Module1"\n\' 日本語テスト\nPublic Sub Hello()\n  MsgBox "こんにちは"\nEnd Sub\n';
		writeFileSync(srcPath, srcContent);
		writeFileSync(manifestPath, toml);

		const { project, dependencies } = await setup(fixtureDir);
		const graph = await loadFromProject(project, dependencies);
		expect(graph.components.length).toBe(1);
		expect(graph.components[0].details.sourceEncoding).toBe("cp932");
	});
});
