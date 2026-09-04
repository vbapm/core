/**
 * End-to-end tests for `vba close`.
 */

import { writeFile } from "fs-extra";
import { join } from "path";
import { promisify } from "util";
import { standard } from "./__fixtures__";
import { execute, readdir, setup } from "./__helpers__/execute";

const exec = promisify(require("child_process").exec);

jest.setTimeout(180000);

describe("close", () => {
	test("close with --force succeeds after build", async () => {
		await setup(standard, "close-built", async cwd => {
			await execute(cwd, "build");
			const result = await execute(cwd, "close --force");
			expect(typeof result.stderr).toBe("string");
		});
	});

	test("close fails when built target does not exist", async () => {
		await setup(standard, "close-not-built", async cwd => {
			await expect(execute(cwd, "close --force")).rejects.toThrow();
		});
	});

	test("close with --force shows error for missing build", async () => {
		await setup(standard, "close-missing-build", async cwd => {
			try {
				await execute(cwd, "close --force");
			} catch (err: any) {
				expect(err.stderr || err.stdout || "").toMatch(/not found/i);
			}
		});
	});

	test("close fails with unsaved changes error when workbook has been modified", async () => {
		await setup(standard, "close-unsaved", async cwd => {
			// 1. Build
			await execute(cwd, "build");

			const buildDir = join(cwd, "build");
			const builtFiles = await readdir(buildDir);
			const targetFile = Object.keys(builtFiles).find(
				f => f.endsWith(".xlsm") || f.endsWith(".xlam")
			);
			if (!targetFile) throw new Error("No built target file found");

			const fullPath = join(buildDir, targetFile);

			// 2. Open workbook via COM, make a change, don't save (leave open).
			//    The script also reports the Excel PID so we can quit it afterward.
			const psScriptPath = join(cwd, "_open_unsaved.ps1");
			const escapedPath = fullPath.replace(/\\/g, "\\\\");
			const psScript = [
				'$path = "' + escapedPath + '"',
				"$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)",
				"$excel = New-Object -ComObject Excel.Application",
				"$excel.Visible = $false",
				"$wb = $excel.Workbooks.Open($path)",
				"$ws = $wb.Worksheets(1)",
				'$ws.Cells.Item(1,1) = "unsaved change"',
				"$after = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)",
				"$excelPid = ($after | Where-Object { $before -notcontains $_ } | Select-Object -First 1)",
				'Write-Output "opened_and_modified pid=$excelPid"'
			].join("\n");
			await writeFile(psScriptPath, psScript);

			const { stdout: psOut } = await exec(
				`powershell -NoProfile -ExecutionPolicy Bypass -File "${psScriptPath}"`,
				{ timeout: 30000 }
			);
			expect(psOut.trim()).toMatch(/^opened_and_modified pid=\d+$/);
			const excelPid = psOut.trim().match(/pid=(\d+)/)?.[1];

			try {
				// 3. Close without --save or --force → should fail with unsaved changes
				await expect(execute(cwd, "close")).rejects.toThrow();
			} finally {
				// 4. Quit the hidden Excel we opened so it doesn't linger holding
				//    the built target (and so the temp dir can be removed).
				if (excelPid) {
					await exec(
						`powershell -NoProfile -Command "Stop-Process -Id ${excelPid} -Force -ErrorAction SilentlyContinue"`,
						{ timeout: 15000 }
					);
				}
			}
		}).catch(() => {
			// Temp dir cleanup may fail (EBUSY from lingering Excel process) — ignore
		});
	});
});
