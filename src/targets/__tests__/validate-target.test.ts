import { mkdtemp, rm } from "fs/promises";
import { join } from "path";
import { tmpdir } from "os";
import { ensureDir, writeFile } from "fs-extra";
import { CliError, ErrorCode } from "../../errors";
import { Target } from "../../manifest/target";
import { validateTarget } from "../validate-target";

const WORKSHEET_REL_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

const target: Target = {
	name: "fixture",
	type: "xlsm",
	path: "",
	filename: "fixture.xlsm"
};

async function writeWorkbook(dir: string, sheets: string): Promise<void> {
	await writeFile(
		join(dir, "xl", "workbook.xml"),
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${sheets}</sheets>
</workbook>`
	);
}

async function writeRels(dir: string, relationships: string): Promise<void> {
	await writeFile(
		join(dir, "xl", "_rels", "workbook.xml.rels"),
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`
	);
}

async function writeWorksheet(dir: string, name: string): Promise<void> {
	await writeFile(
		join(dir, "xl", "worksheets", name),
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`
	);
}

describe("validateTarget", () => {
	let dir: string;

	beforeEach(async () => {
		dir = await mkdtemp(join(tmpdir(), "vbapm-validate-target-"));
		target.path = dir;
		await ensureDir(join(dir, "xl", "worksheets"));
		await ensureDir(join(dir, "xl", "_rels"));
	});

	afterEach(async () => {
		await rm(dir, { recursive: true, force: true });
	});

	test("accepts a valid target with a worksheet relationship", async () => {
		await writeWorkbook(dir, '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>');
		await writeRels(
			dir,
			`\n  <Relationship Id="rId1" Type="${WORKSHEET_REL_TYPE}" Target="worksheets/sheet1.xml"/>`
		);
		await writeWorksheet(dir, "sheet1.xml");

		await expect(validateTarget(target, dir)).resolves.toBeUndefined();
	});

	test("rejects a sheet without an r:id", async () => {
		await writeWorkbook(dir, '<sheet name="Module2" sheetId="5" state="veryHidden"/>');
		await writeRels(dir, "");

		await expect(validateTarget(target, dir)).rejects.toMatchObject({
			code: ErrorCode.TargetInvalid,
			message: expect.stringContaining('Worksheet "Module2"')
		});
	});

	test("rejects a sheet whose r:id has no matching relationship", async () => {
		await writeWorkbook(dir, '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>');
		await writeRels(dir, "");

		await expect(validateTarget(target, dir)).rejects.toMatchObject({
			code: ErrorCode.TargetInvalid,
			message: expect.stringContaining('references relationship "rId1"')
		});
	});

	test("rejects a sheet whose relationship target file does not exist", async () => {
		await writeWorkbook(dir, '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>');
		await writeRels(
			dir,
			`\n  <Relationship Id="rId1" Type="${WORKSHEET_REL_TYPE}" Target="worksheets/sheet1.xml"/>`
		);
		// no sheet1.xml written

		await expect(validateTarget(target, dir)).rejects.toMatchObject({
			code: ErrorCode.TargetInvalid,
			message: expect.stringContaining("does not")
		});
	});

	test("rejects an empty r:id even when a worksheet relationship exists", async () => {
		// The EmailManager case: r:id="" + veryHidden, no backing part.
		await writeWorkbook(
			dir,
			'<sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Module2" sheetId="5" state="veryHidden" r:id=""/>'
		);
		await writeRels(
			dir,
			`\n  <Relationship Id="rId1" Type="${WORKSHEET_REL_TYPE}" Target="worksheets/sheet1.xml"/>`
		);
		await writeWorksheet(dir, "sheet1.xml");

		await expect(validateTarget(target, dir)).rejects.toMatchObject({
			code: ErrorCode.TargetInvalid,
			message: expect.stringContaining('Worksheet "Module2"')
		});
	});
});
