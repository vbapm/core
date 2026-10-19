/**
 * End-to-end test for `normalize-worksheet-names` (export-side worksheet renaming).
 */

import { pathExists, readFile, remove, writeFile } from "fs-extra";
import { join } from "path";
import { withDrawing } from "./__fixtures__";
import { execute, setup } from "./__helpers__/execute";
import { unzip, zip } from "../src/utils/zip";

jest.setTimeout(180000);

describe("normalize-worksheet-names", () => {
	test("renames worksheet sidecar .rels file alongside worksheet XML", async () => {
		await setup(withDrawing, "normalize-sidecar-rels", async cwd => {
			// Build creates the xlsm from targets/xlsm OOXML (which contains a drawing)
			await execute(cwd, "build");

			// Export runs normalizeWorksheetNames: sheet1.xml → shtSheet1.xml
			// The sidecar _rels/sheet1.xml.rels must also be renamed → shtSheet1.xml.rels
			await execute(cwd, "export --target xlsm");

			const relsDir = join(cwd, "targets/xlsm/xl/worksheets/_rels");
			expect(await pathExists(join(relsDir, "shtSheet1.xml.rels"))).toBe(true);
			expect(await pathExists(join(relsDir, "sheet1.xml.rels"))).toBe(false);
		});
	});

	test("xml-only extract normalizes a sheet without a codeName in its XML", async () => {
		await setup(withDrawing, "normalize-xml-only", async cwd => {
			const source = join(cwd, "src", "Sheet1.cls");
			const originalSource = await readFile(source);
			await execute(cwd, "build");
			const workbook = join(cwd, "build", "with-drawing.xlsm");
			const unpacked = join(cwd, "unpacked");
			await unzip(workbook, unpacked);
			const sheetPath = join(unpacked, "xl", "worksheets", "sheet1.xml");
			const sheet = await readFile(sheetPath, "utf8");
			await writeFile(sheetPath, sheet.replace(' codeName="Sheet1"', ""));
			await zip(unpacked, workbook);
			await remove(unpacked);

			await execute(cwd, "extract --xml-only --background");

			const worksheets = join(cwd, "targets", "xlsm", "xl", "worksheets");
			expect(await pathExists(join(worksheets, "shtSheet1.xml"))).toBe(true);
			expect(await pathExists(join(worksheets, "sheet1.xml"))).toBe(false);
			expect(await readFile(source)).toEqual(originalSource);
		});
	});
});
