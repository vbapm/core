import { mkdir, mkdtemp, readFile as readFileNative, rm, writeFile as writeFileNative } from "fs/promises";
import { join } from "path";
import { tmpdir } from "os";
import { readFile } from "../../../utils/fs";
import { findElementByName, parseXml } from "../../../utils/xml";
import {
	preserveWorksheetSelections,
	readPreviousWorksheetData
} from "../preserve-worksheet-selection";

const OLD_XML = `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr codeName="Sheet1"/>
  <dimension ref="A1"/>
  <sheetViews><sheetView workbookViewId="0"><selection activeCell="G14" sqref="G14"/></sheetView></sheetViews>
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`;

function worksheetXml(activeCell: string, sqref: string, dimension = "A1") {
	return OLD_XML.replace(
		'activeCell="G14" sqref="G14"',
		`activeCell="${activeCell}" sqref="${sqref}"`
	).replace('dimension ref="A1"', `dimension ref="${dimension}"`);
}

describe("preserveWorksheetSelections", () => {
	let tmp: string;

	beforeEach(async () => {
		tmp = await mkdtemp(join(tmpdir(), "vbapm-selection-"));
		await mkdir(join(tmp, "old/xl/worksheets"), { recursive: true });
		await mkdir(join(tmp, "new/xl/worksheets"), { recursive: true });
	});

	afterEach(async () => {
		await rm(tmp, { recursive: true, force: true });
	});

	test("restores selection attributes when they are the only worksheet XML changes", async () => {
		await writeFileNative(join(tmp, "old/xl/worksheets/shtSheet1.xml"), OLD_XML);
		await writeFileNative(join(tmp, "new/xl/worksheets/sheet1.xml"), worksheetXml("A1", "A1"));

		const previous = await readPreviousWorksheetData(join(tmp, "old"));
		await preserveWorksheetSelections(join(tmp, "new"), previous);

		const xml = parseXml(await readFile(join(tmp, "new/xl/worksheets/sheet1.xml")));
		const worksheet = findElementByName(xml.elements, "worksheet");
		const sheetViews = findElementByName(worksheet?.elements, "sheetViews");
		const sheetView = findElementByName(sheetViews?.elements, "sheetView");
		const selection = findElementByName(sheetView?.elements, "selection");
		expect(selection?.attributes).toMatchObject({ activeCell: "G14", sqref: "G14" });
	});

	test("keeps selection changes when worksheet data also changed", async () => {
		await writeFileNative(join(tmp, "old/xl/worksheets/sheet1.xml"), OLD_XML);
		await writeFileNative(join(tmp, "new/xl/worksheets/sheet1.xml"), worksheetXml("A1", "A1", "B2"));

		const previous = await readPreviousWorksheetData(join(tmp, "old"));
		await preserveWorksheetSelections(join(tmp, "new"), previous);

		const xml = parseXml(await readFile(join(tmp, "new/xl/worksheets/sheet1.xml")));
		const worksheet = findElementByName(xml.elements, "worksheet");
		const sheetViews = findElementByName(worksheet?.elements, "sheetViews");
		const sheetView = findElementByName(sheetViews?.elements, "sheetView");
		const selection = findElementByName(sheetView?.elements, "selection");
		expect(selection?.attributes).toMatchObject({ activeCell: "A1", sqref: "A1" });
	});
});
