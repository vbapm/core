import { mkdtemp, mkdir, rm, writeFile } from "fs/promises";
import { tmpdir } from "os";
import { join } from "path";
import { validateExportTarget } from "../addin";
import { Target } from "../manifest/target";

const target: Target = {
	name: "ProjectName",
	type: "xlsm",
	path: "target",
	filename: "ProjectName.xlsm"
};

describe("validateExportTarget", () => {
	let temporaryDirectory: string;

	beforeEach(async () => {
		temporaryDirectory = await mkdtemp(join(tmpdir(), "vbapm-export-target-"));
	});

	afterEach(async () => {
		await rm(temporaryDirectory, { recursive: true, force: true });
	});

	test("accepts an existing target file", async () => {
		const file = join(temporaryDirectory, target.filename);
		await writeFile(file, "");

		await expect(validateExportTarget(file, target, "ProjectName")).resolves.toBeUndefined();
	});

	test("reports the first missing directory", async () => {
		const existing = join(temporaryDirectory, "existing");
		const missing = join(existing, "missing");
		await mkdir(existing);

		await expect(
			validateExportTarget(join(missing, "nested", target.filename), target, "ProjectName")
		).rejects.toThrow(`The directory containing the target file does not exist:\n\n  "${missing}"`);
	});

	test("suggests the only workbook with a matching extension", async () => {
		await writeFile(join(temporaryDirectory, "Master - Billing.xlsm"), "");
		await writeFile(join(temporaryDirectory, "Notes.xlsx"), "");

		await expect(
			validateExportTarget(join(temporaryDirectory, target.filename), target, "ProjectName")
		).rejects.toThrow(
			`Found other .xlsm files in "${temporaryDirectory}":\n` +
				`  - "Master - Billing.xlsm"\n\n` +
				`The target filename defaults to the project name "ProjectName". ` +
				`If the workbook name differs, set target.name in vbaproject.toml without the extension:\n\n` +
				`  [project]\n  target = { type = "xlsm", name = "Master - Billing" }`
		);
	});

	test("lists multiple matching workbooks without choosing one", async () => {
		await writeFile(join(temporaryDirectory, "First.xlsm"), "");
		await writeFile(join(temporaryDirectory, "Second.xlsm"), "");

		await expect(
			validateExportTarget(join(temporaryDirectory, target.filename), target, "ProjectName")
		).rejects.toThrow(
			`  - "First.xlsm"\n` +
				`  - "Second.xlsm"\n\n` +
				`The target filename defaults to the project name "ProjectName". ` +
				`If the workbook name differs, set target.name in vbaproject.toml without the extension:\n\n` +
				`  [project]\n  target = { type = "xlsm", name = "WORKBOOK_NAME" }`
		);
	});

	test("does not suggest target.name when no matching files exist", async () => {
		await expect(
			validateExportTarget(join(temporaryDirectory, target.filename), target, "ProjectName")
		).rejects.toThrow(`The target file does not exist:`);

		await expect(
			validateExportTarget(join(temporaryDirectory, target.filename), target, "ProjectName")
		).rejects.not.toThrow(`target.name`);
	});
});
