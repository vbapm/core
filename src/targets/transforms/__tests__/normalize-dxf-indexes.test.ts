import { mkdtemp, rm } from "fs/promises";
import { join } from "path";
import { tmpdir } from "os";
import { readFile } from "../../../utils/fs";
import { ensureDir, writeFile } from "fs-extra";
import { Element } from "xml-js";
import { parseXml } from "../../../utils/xml";
import { normalizeDxfIndexes } from "../normalize-dxf-indexes";

function findElementsByName(elements: Element[] | undefined, name: string): Element[] {
	return (elements ?? []).flatMap(element => [
		...((element.name ?? "").split(":").pop() === name ? [element] : []),
		...findElementsByName(element.elements, name)
	]);
}

describe("normalizeDxfIndexes", () => {
	let tmp: string;
	let current: string;
	let previous: string;

	beforeEach(async () => {
		tmp = await mkdtemp(join(tmpdir(), "vbapm-dxf-normalize-"));
		current = join(tmp, "current");
		previous = join(tmp, "previous");
		await ensureDir(join(current, "xl", "worksheets"));
		await ensureDir(join(current, "xl", "tables"));
		await ensureDir(join(previous, "xl"));
	});

	afterEach(async () => {
		await rm(tmp, { recursive: true, force: true });
	});

	async function writeStyles(root: string, dxfs: string[]) {
		await ensureDir(join(root, "xl"));
		await writeFile(
			join(root, "xl", "styles.xml"),
			`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dxfs count="${dxfs.length}">${dxfs.join("")}</dxfs></styleSheet>`
		);
	}

	async function readDxfs(root: string): Promise<string[]> {
		const xml = parseXml(await readFile(join(root, "xl", "styles.xml")));
		const stylesheet = findElementsByName(xml.elements, "styleSheet")[0];
		const dxfs = findElementsByName(stylesheet?.elements, "dxfs")[0];
		return (dxfs?.elements ?? [])
			.filter(element => (element.name ?? "").split(":").pop() === "dxf")
			.map(element => JSON.stringify(element));
	}

	test("restores reordered DXFs and remaps worksheet and table references", async () => {
		const red = '<dxf><font><color rgb="FFFF0000"/></font></dxf>';
		const blue = '<dxf><font><color rgb="FF0000FF"/></font></dxf>';
		await writeStyles(previous, [red, blue]);
		await writeStyles(current, [blue, red]);
		await writeFile(
			join(current, "xl", "worksheets", "sheet1.xml"),
			'<worksheet><conditionalFormatting><cfRule dxfId="0"/><cfRule dxfId="1"/></conditionalFormatting></worksheet>'
		);
		await writeFile(
			join(current, "xl", "tables", "table1.xml"),
			'<table><tableColumn dataDxfId="1" headerRowDxfId="0" totalsRowDxfId="1"/></table>'
		);
		await normalizeDxfIndexes(current, previous);

		expect(await readDxfs(current)).toEqual(await readDxfs(previous));
		const worksheet = parseXml(await readFile(join(current, "xl", "worksheets", "sheet1.xml")));
		const rules = findElementsByName(worksheet.elements, "cfRule");
		expect(rules.map(rule => rule.attributes?.dxfId)).toEqual(["1", "0"]);
		const table = parseXml(await readFile(join(current, "xl", "tables", "table1.xml")));
		const column = findElementsByName(table.elements, "tableColumn")[0];
		expect(column?.attributes).toMatchObject({
			dataDxfId: "0",
			headerRowDxfId: "1",
			totalsRowDxfId: "0"
		});
	});

	test("keeps removed definitions in place and appends new definitions", async () => {
		const a = "<dxf><font><b/></font></dxf>";
		const b = "<dxf><font><i/></font></dxf>";
		const c = "<dxf><font><u/></font></dxf>";
		const added = "<dxf><font><strike/></font></dxf>";
		await writeStyles(previous, [a, b, c]);
		await writeStyles(current, [c, a, added]);
		await writeFile(join(current, "xl", "tables", "table1.xml"), '<table dataDxfId="2"/>');

		await normalizeDxfIndexes(current, previous);

		expect((await readDxfs(current)).length).toBe(4);
		const xml = parseXml(await readFile(join(current, "xl", "tables", "table1.xml")));
		expect(findElementsByName(xml.elements, "table")[0]?.attributes?.dataDxfId).toBe("3");
	});

	test("treats prefix and attribute-order changes as equivalent", async () => {
		await writeFile(
			join(previous, "xl", "styles.xml"),
			'<styleSheet xmlns="urn:styles"><dxfs count="1"><dxf><font><color rgb="FF112233" theme="1"/></font></dxf></dxfs></styleSheet>'
		);
		await writeFile(
			join(current, "xl", "styles.xml"),
			'<s:styleSheet xmlns:s="urn:styles"><s:dxfs count="1"><s:dxf><s:font><s:color theme="1" rgb="FF112233"/></s:font></s:dxf></s:dxfs></s:styleSheet>'
		);

		await normalizeDxfIndexes(current, previous);

		expect((await readDxfs(current)).length).toBe(1);
	});

	test("appends a modified definition and maps its reference to the new index", async () => {
		await writeStyles(previous, ["<dxf><font><b/></font></dxf>"]);
		await writeStyles(current, ["<dxf><font><i/></font></dxf>"]);
		await writeFile(join(current, "xl", "worksheets", "sheet1.xml"), '<worksheet dxfId="0"/>');

		await normalizeDxfIndexes(current, previous);

		expect((await readDxfs(current)).length).toBe(2);
		const xml = parseXml(await readFile(join(current, "xl", "worksheets", "sheet1.xml")));
		expect(findElementsByName(xml.elements, "worksheet")[0]?.attributes?.dxfId).toBe("1");
	});

	test("matches duplicate DXFs deterministically and is idempotent", async () => {
		const duplicate = "<dxf><font><b/></font></dxf>";
		const other = "<dxf><font><i/></font></dxf>";
		await writeStyles(previous, [duplicate, other, duplicate]);
		await writeStyles(current, [duplicate, duplicate, other]);
		await writeFile(join(current, "xl", "tables", "table1.xml"), '<table dataDxfId="2"/>');

		await normalizeDxfIndexes(current, previous);
		const firstPass = await readFile(join(current, "xl", "styles.xml"));
		await normalizeDxfIndexes(current, previous);

		expect(await readFile(join(current, "xl", "styles.xml"))).toEqual(firstPass);
		const xml = parseXml(await readFile(join(current, "xl", "tables", "table1.xml")));
		expect(findElementsByName(xml.elements, "table")[0]?.attributes?.dataDxfId).toBe("1");
	});

	test("leaves malformed and out-of-range references unchanged", async () => {
		await writeStyles(previous, ["<dxf><font><b/></font></dxf>"]);
		await writeStyles(current, ["<dxf><font><b/></font></dxf>"]);
		const source = '<worksheet dxfId="bad"><conditionalFormatting dxfId="4"/></worksheet>';
		await writeFile(join(current, "xl", "worksheets", "sheet1.xml"), source);

		await expect(normalizeDxfIndexes(current, previous)).resolves.toBeUndefined();
		const output = await readFile(join(current, "xl", "worksheets", "sheet1.xml"), "utf8");
		expect(output).toContain('dxfId="bad"');
		expect(output).toContain('dxfId="4"');
	});

	test("does nothing when styles are missing", async () => {
		await expect(normalizeDxfIndexes(current, previous)).resolves.toBeUndefined();
	});
});
