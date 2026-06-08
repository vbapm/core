import { mkdtemp, rm } from "fs/promises";
import { join } from "path";
import { tmpdir } from "os";
import { copy } from "fs-extra";
import { pathExists, readFile } from "../../../utils/fs";
import { parseXml, findElement, findElementByName } from "../../../utils/xml";
import { normalizeWorksheetNames } from "../normalize-worksheet-names";

const FIXTURES = join(__dirname, "__fixtures__", "normalize-worksheet-names");
const WS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const WORKSHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

describe("normalizeWorksheetNames", () => {
	let tmp: string;

	beforeEach(async () => {
		tmp = await mkdtemp(join(tmpdir(), "vbapm-normalize-"));
	});

	afterEach(async () => {
		await rm(tmp, { recursive: true, force: true });
	});

	async function copyFixture(name: string) {
		await copy(join(FIXTURES, name), tmp);
	}

	async function fileExists(relativePath: string): Promise<boolean> {
		return pathExists(join(tmp, relativePath));
	}

	async function readContents(relativePath: string): Promise<string> {
		return (await readFile(join(tmp, relativePath))).toString("utf8");
	}

	test("renames sheet1.xml to sht{codeName}.xml and updates xml/rels", async () => {
		await copyFixture("basic");
		await normalizeWorksheetNames(tmp);

		expect(await fileExists("xl/worksheets/sheet1.xml")).toBe(false);
		expect(await fileExists("xl/worksheets/shtSheet1.xml")).toBe(true);

		const rels = parseXml(await readContents("xl/_rels/workbook.xml.rels"));
		const wsRel = findElement(
			findElementByName(rels.elements, "Relationships")!.elements,
			el => el.name === "Relationship" && el.attributes?.Type === WS_REL_TYPE
		);
		expect(wsRel?.attributes?.Target).toBe("worksheets/shtSheet1.xml");

		const ct = parseXml(await readContents("[Content_Types].xml"));
		const types = findElementByName(ct.elements, "Types")!;
		expect(
			findElement(
				types.elements,
				el => el.name === "Override" && el.attributes?.PartName === "/xl/worksheets/shtSheet1.xml"
			)?.attributes?.ContentType
		).toBe(WORKSHEET_CT);
		expect(
			findElement(
				types.elements,
				el => el.name === "Override" && el.attributes?.PartName === "/xl/worksheets/sheet1.xml"
			)
		).toBeUndefined();
	});

	// Excel always assigns a codeName, but this guards against
	// manually-edited OOXML or third-party generated files.
	test("keeps original name when sheet has no codeName", async () => {
		await copyFixture("no-codename");
		await normalizeWorksheetNames(tmp);

		expect(await fileExists("xl/worksheets/sheet1.xml")).toBe(true);
		expect(await fileExists("xl/worksheets/shtSheet1.xml")).toBe(false);
	});

	test("no-op when worksheet already has the sht{codeName}.xml name", async () => {
		await copyFixture("already-correct");
		await normalizeWorksheetNames(tmp);

		expect(await fileExists("xl/worksheets/shtSheet1.xml")).toBe(true);
		expect(await readContents("[Content_Types].xml")).toContain("/xl/worksheets/shtSheet1.xml");
	});

	test("renames sidecar _rels/<sheet>.rels alongside worksheet", async () => {
		await copyFixture("with-sidecar");
		await normalizeWorksheetNames(tmp);

		expect(await fileExists("xl/worksheets/shtSheet1.xml")).toBe(true);
		expect(await fileExists("xl/worksheets/sheet1.xml")).toBe(false);
		expect(await fileExists("xl/worksheets/_rels/shtSheet1.xml.rels")).toBe(true);
		expect(await fileExists("xl/worksheets/_rels/sheet1.xml.rels")).toBe(false);
		expect(await readContents("xl/worksheets/_rels/shtSheet1.xml.rels")).toContain("drawing1.xml");
	});

	test("skips gracefully when workbook.xml.rels is missing", async () => {
		await expect(normalizeWorksheetNames(tmp)).resolves.toBeUndefined();
	});

	test("skips when workbook.xml.rels has no worksheet relationships", async () => {
		await copyFixture("no-worksheet-rels");
		await normalizeWorksheetNames(tmp);

		expect(await fileExists("xl/worksheets/sheet1.xml")).toBe(true);
	});

	test("renames multiple worksheets correctly", async () => {
		await copyFixture("multiple");
		await normalizeWorksheetNames(tmp);

		expect(await fileExists("xl/worksheets/shtSheet1.xml")).toBe(true);
		expect(await fileExists("xl/worksheets/shtDashboard.xml")).toBe(true);
		expect(await fileExists("xl/worksheets/sheet1.xml")).toBe(false);
		expect(await fileExists("xl/worksheets/sheet2.xml")).toBe(false);

		expect(await readContents("xl/_rels/workbook.xml.rels")).toContain("worksheets/shtSheet1.xml");
		expect(await readContents("xl/_rels/workbook.xml.rels")).toContain(
			"worksheets/shtDashboard.xml"
		);
		expect(await readContents("[Content_Types].xml")).toContain("/xl/worksheets/shtSheet1.xml");
		expect(await readContents("[Content_Types].xml")).toContain("/xl/worksheets/shtDashboard.xml");
	});

	test("renames worksheets even when [Content_Types].xml is missing", async () => {
		await copyFixture("basic");
		await rm(join(tmp, "[Content_Types].xml"));

		await normalizeWorksheetNames(tmp);

		expect(await fileExists("xl/worksheets/shtSheet1.xml")).toBe(true);
		expect(await fileExists("xl/worksheets/sheet1.xml")).toBe(false);
	});

	test("removes stale duplicate when destination already exists", async () => {
		await copyFixture("stale-dest");
		await normalizeWorksheetNames(tmp);

		expect(await fileExists("xl/worksheets/shtSheet1.xml")).toBe(true);
	});
});
