import { loadFromProject } from "../load-from-project";
import { reset, setup } from "../../../tests/__helpers__/project";
import { dir } from "../../../tests/__fixtures__";
import { join } from "../../utils/path";
import { writeFileSync, mkdirSync, rmSync } from "fs";
import dedent from "@timhall/dedent";
import jschardet from "jschardet";

afterEach(() => {
	reset();
	jest.restoreAllMocks();
});

describe("encoding validation", () => {
	const fixtureDir = join(dir, "projects", "encoding-test");

	// Fixture directory for encoding tests
	beforeAll(() => {
		mkdirSync(join(fixtureDir, "src"), { recursive: true });
		mkdirSync(join(fixtureDir, "targets"), { recursive: true });
	});

	// Clean up fixture directory after all tests
	afterAll(() => {
		rmSync(fixtureDir, { recursive: true, force: true });
	});

	test("ASCII-only source without encoding → builds fine", async () => {
		const toml = dedent`
		[package]
		name = "ascii-only"
		version = "1.0.0"
		authors = ["Test"]

		[src]
		Module1 = "src/Module1.bas"
		`;

		const srcPath = join(fixtureDir, "src", "Module1.bas");
		const manifestPath = join(fixtureDir, "vbaproject.toml");

		await writeFileSync(
			srcPath,
			'Attribute VB_Name = "Module1"\nPublic Sub Hello()\n  MsgBox "Hi"\nEnd Sub\n'
		);
		await writeFileSync(manifestPath, toml);

		const { project, dependencies } = await setup(fixtureDir);
		const graph = await loadFromProject(project, dependencies);
		expect(graph.components.length).toBe(1);
		expect(graph.components[0].code).toContain("Hello");
	});

	const nonAsciiSource = dedent`
		Attribute VB_Name = "Module1"
		' Déjà vu – naïve façade
		Public Sub Hello()
		  MsgBox "é"  ' café
		End Sub
	`;
	// Replace \n with \r\n for Windows-style line endings
	const nonAsciiSourceWindows = nonAsciiSource.replace(/\n/g, "\r\n");
	const utf16leSource = Buffer.from(nonAsciiSourceWindows, "utf16le");

	// Array<[description, expected encoding, source buffer]>
	const certainBomCases: Array<[string, string, Buffer]> = [
		[
			"UTF-8 BOM",
			"utf-8",
			Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(nonAsciiSource, "utf8")])
		],
		[	"UTF-16 LE BOM", 
			"utf-16le", 
			Buffer.concat([Buffer.from([0xff, 0xfe]), utf16leSource])],
		[
			"UTF-16 BE BOM",
			"utf-16be",
			Buffer.concat([Buffer.from([0xfe, 0xff]), Buffer.from(utf16leSource).swap16()])
		]
	];

	test.each(certainBomCases)(
		"non-ASCII %s without encoding uses certain %s detection",
		async (_name, expectedEncoding, sourceBuffer) => {
			const toml = dedent`
			[package]
			name = "non-ascii-no-enc"
			version = "1.0.0"
			authors = ["Test"]

			[src]
			Module1 = "src/Module1.bas"
			`;

			const srcPath = join(fixtureDir, "src", "Module1.bas");
			const manifestPath = join(fixtureDir, "vbaproject.toml");

			writeFileSync(srcPath, sourceBuffer);
			writeFileSync(manifestPath, toml);

			const detectAllSpy = jest.spyOn(jschardet, "detectAll");
			const { project, dependencies } = await setup(fixtureDir);
			// Attempting to load the project should throw due 
			// to missing encoding from the manifest
			await expect(loadFromProject(project, dependencies)).rejects.toThrow(
				`encoding = "${expectedEncoding}"`
			);
			// Making sure that jschardet.detectAll was not called 
			// since the encoding should be certain from the BOM
			expect(detectAllSpy).not.toHaveBeenCalled();
		}
	);

	test("non-ASCII with encoding in [source] → builds fine", async () => {
		const toml = dedent`
		[package]
		name = "non-ascii-enc"
		version = "1.0.0"
		authors = ["Test"]

		[source]
		encoding = "cp1252"

		[src]
		Module1 = "src/Module1.bas"
		`;

		const srcPath = join(fixtureDir, "src", "Module1.bas");
		const manifestPath = join(fixtureDir, "vbaproject.toml");

		const srcContent =
			'Attribute VB_Name = "Module1"\n\' Déjà vu – naïve façade\nPublic Sub Hello()\n  MsgBox "é"  \' café\nEnd Sub\n';
		writeFileSync(srcPath, srcContent);
		writeFileSync(manifestPath, toml);

		const { project, dependencies } = await setup(fixtureDir);
		const graph = await loadFromProject(project, dependencies);
		expect(graph.components.length).toBe(1);
	});

	test("per-source encoding overrides project-level", async () => {
		const toml = dedent`
		[package]
		name = "per-source-enc"
		version = "1.0.0"
		authors = ["Test"]

		[source]
		encoding = "cp1252"

		[src]
		Module1 = { path = "src/Module1.bas", encoding = "cp932" }
		`;

		const srcPath = join(fixtureDir, "src", "Module1.bas");
		const manifestPath = join(fixtureDir, "vbaproject.toml");

		// Japanese characters (CP932)
		const srcContent =
			'Attribute VB_Name = "Module1"\n\' 日本語テスト\nPublic Sub Hello()\n  MsgBox "こんにちは"\nEnd Sub\n';
		writeFileSync(srcPath, srcContent);
		writeFileSync(manifestPath, toml);

		const { project, dependencies } = await setup(fixtureDir);
		const graph = await loadFromProject(project, dependencies);
		expect(graph.components.length).toBe(1);
		expect(graph.components[0].details.sourceEncoding).toBe("cp932");
	});
});
