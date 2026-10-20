import { readFile, remove, writeFile } from "fs-extra";
import { join } from "path";
import { withDrawing } from "./__fixtures__";
import { execute, setup } from "./__helpers__/execute";
import { unzip, zip } from "../src/utils/zip";

jest.setTimeout(180000);

describe("normalize-dxf-indexes", () => {
	test("restores stable DXF indexes and rewrites references during XML-only extraction", async () => {
		await setup(withDrawing, "normalize-dxf-indexes", async cwd => {
			const target = join(cwd, "targets", "xlsm");
			const stylesPath = join(target, "xl", "styles.xml");
			const red = '<dxf><font><color rgb="FFFF0000"/></font></dxf>';
			const blue = '<dxf><font><color rgb="FF0000FF"/></font></dxf>';
			await execute(cwd, "build");
			const originalStyles = await readFile(stylesPath, "utf8");
			await writeFile(
				stylesPath,
				originalStyles.replace('<dxfs count="0"/>', `<dxfs count="2">${red}${blue}</dxfs>`)
			);

			const workbook = join(cwd, "build", "with-drawing.xlsm");
			const unpacked = join(cwd, "unpacked");
			await unzip(workbook, unpacked);
			const workbookStylesPath = join(unpacked, "xl", "styles.xml");
			const workbookStyles = await readFile(workbookStylesPath, "utf8");
			await writeFile(
				workbookStylesPath,
				workbookStyles.replace('<dxfs count="0"/>', `<dxfs count="2">${blue}${red}</dxfs>`)
			);

			const worksheetPath = join(unpacked, "xl", "worksheets", "sheet1.xml");
			const worksheet = await readFile(worksheetPath, "utf8");
			await writeFile(
				worksheetPath,
				worksheet.replace(
					"</worksheet>",
					'<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1"><formula>1=1</formula></cfRule></conditionalFormatting></worksheet>'
				)
			);
			await zip(unpacked, workbook);
			await remove(unpacked);

			await execute(cwd, "extract --xml-only --skip-sheet-name-normalization --background");

			const normalizedStyles = await readFile(stylesPath, "utf8");
			expect(normalizedStyles.indexOf('rgb="FFFF0000"')).toBeLessThan(
				normalizedStyles.indexOf('rgb="FF0000FF"')
			);
			const normalizedWorksheet = await readFile(
				join(target, "xl", "worksheets", "sheet1.xml"),
				"utf8"
			);
			expect(normalizedWorksheet).toContain('dxfId="1"');
		});
	});
});
