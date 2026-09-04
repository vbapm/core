/**
 * End-to-end test for `normalize-worksheet-names` (export-side worksheet renaming).
 */

import { pathExists } from "fs-extra";
import { join } from "path";
import { withDrawing } from "./__fixtures__";
import { execute, setup } from "./__helpers__/execute";

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
});
