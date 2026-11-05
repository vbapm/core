/**
 * End-to-end repro for an orphaned worksheet declaration in a target.
 *
 * ## Summary
 *
 * Building a project whose target template declares a worksheet with an empty
 * relationship id (`r:id=""`) must fail with a clear `target-invalid` error
 * rather than a cryptic Excel `1004: Save method of Workbook class failed`.
 *
 * See `docs/development/veryhidden-sheet-corruption-issue.md` for the full
 * investigation. In short: the malformed OOXML declares an orphaned `<sheet>`
 * with no backing worksheet part. Previously Excel would silently drop it on
 * open and `state="veryHidden"` would surface a `1004` on save; now vbapm
 * validates the worksheet graph up front and reports the problem clearly.
 *
 * ## Fixture
 *
 * `orphaned-sheet` is a copy of `with-drawing` whose `workbook.xml` declares an
 * extra sheet:
 *
 * ```xml
 * <sheet name="Module2" sheetId="5" state="veryHidden" r:id=""/>
 * ```
 *
 * The empty `r:id` means the sheet has no backing worksheet part or
 * relationship, and no content-type override.
 *
 * ## Repro assertion
 *
 * The test asserts that `build` fails with the validation error (mentioning the
 * orphaned worksheet), and that no workbook is produced.
 */

import { pathExists } from "fs-extra";
import { join } from "path";
import { orphanedSheet } from "./__fixtures__";
import { execute, setup } from "./__helpers__/execute";

jest.setTimeout(180000);

describe("orphaned worksheet", () => {
	test("build fails with target-invalid when the target has an orphaned worksheet", async () => {
		await setup(orphanedSheet, "build-orphaned-sheet", async cwd => {
			await expect(execute(cwd, "build")).rejects.toMatchObject({
				stderr: expect.stringContaining('Target "orphaned-sheet" is invalid')
			});

			// The orphaned worksheet should never produce a workbook.
			expect(await pathExists(join(cwd, "build", "orphaned-sheet.xlsm"))).toBe(false);
		});
	});
});
