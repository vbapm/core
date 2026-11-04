import { mkdtemp, mkdir, readFile, readdir, rm, writeFile } from "fs/promises";
import { tmpdir } from "os";
import { join } from "path";
import { resolveSourceFiles } from "../../build/apply-extract";
import { writeFile as writeManifestFile } from "../../utils/fs";
import { Project } from "../../project";
import { exportPowerQueries } from "../export-power-queries";

test("syncs Power Query files under the configured source folder", async () => {
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-"));
	const staging = join(root, "staging");
	const sourceQueries = join(root, "custom-src", "Power Queries");
	const stagedQueries = join(staging, "pq");

	try {
		await mkdir(sourceQueries, { recursive: true });
		await mkdir(stagedQueries, { recursive: true });
		await writeFile(join(sourceQueries, "Old.pq"), "old query", "utf8");
		await writeFile(join(sourceQueries, "notes.txt"), "keep", "utf8");
		await writeFile(join(stagedQueries, "New.pq"), 'let Name = "caf\u00e9" in Name', "utf8");

		const project = {
			paths: { dir: root },
			manifest: {
				type: "project",
				name: "power-query-test",
				version: "0.0.0",
				metadata: { authors: [] },
				src: [],
				srcProperties: {
					folder: "custom-src",
					subfolders: { Modules: "Modules", Classes: "Class Modules" }
				},
				dependencies: [],
				references: [],
				devSrc: [],
				devDependencies: [],
				devReferences: []
			}
		} as unknown as Project;
		jest.clearAllMocks();

		await exportPowerQueries(project, staging);

		expect((await readdir(sourceQueries)).sort()).toEqual(["New.pq", "notes.txt"]);
		expect(await readFile(join(sourceQueries, "New.pq"), "utf8")).toContain("caf\u00e9");
		expect(project.manifest.src).toEqual([
			{ name: "PowerQueries", path: join(root, "custom-src", "Power Queries", "*.pq") }
		]);
		expect(project.manifest.srcProperties?.subfolders?.PowerQueries).toBe("Power Queries");
		expect(writeManifestFile).toHaveBeenCalledWith(
			expect.stringContaining("vbaproject.toml"),
			expect.stringContaining('PowerQueries = "custom-src/Power Queries/*.pq"')
		);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});

test("does not load Power Query wildcard matches as VBA components", async () => {
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-"));
	const sourceDirectory = join(root, "src", "Power Queries");

	try {
		await mkdir(sourceDirectory, { recursive: true });
		await writeFile(join(sourceDirectory, "Query.pq"), "not a VBA component", "utf8");

		const project = {
			paths: { dir: root },
			manifest: {
				src: [{ name: "PowerQueries", path: "src/Power Queries/*.pq" }],
				srcProperties: { subfolders: { PowerQueries: "Power Queries" } }
			}
		} as unknown as Project;

		await expect(resolveSourceFiles(project)).resolves.toEqual(new Map());
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});

test("does not sync when query export was not staged", async () => {
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-"));
	const staging = join(root, "staging");
	const sourceQueries = join(root, "src", "pq");

	try {
		await mkdir(sourceQueries, { recursive: true });
		await writeFile(join(sourceQueries, "Existing.pq"), "keep", "utf8");

		const project = {
			paths: { dir: root },
			manifest: { srcProperties: undefined }
		} as unknown as Project;

		await exportPowerQueries(project, staging);

		expect(await readdir(sourceQueries)).toEqual(["Existing.pq"]);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});

test("uses the fallback folder without adding a manifest group when subfolders are unset", async () => {
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-"));
	const staging = join(root, "staging");
	const stagedQueries = join(staging, "pq");
	const fallbackQueries = join(root, "src", "pq");
	const project = {
		paths: { dir: root },
		manifest: { src: [], srcProperties: undefined }
	} as unknown as Project;

	try {
		await mkdir(stagedQueries, { recursive: true });
		await writeFile(join(stagedQueries, "Query.pq"), "query", "utf8");

		await exportPowerQueries(project, staging);

		expect(await readdir(fallbackQueries)).toEqual(["Query.pq"]);
		expect(project.manifest.src).toEqual([]);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});

test("uses an existing Power Query source path without adding a generic folder", async () => {
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-"));
	const staging = join(root, "staging");
	const stagedQueries = join(staging, "pq");
	const sourceQueries = join(root, "src", "EmailManager.xlsm", "Power Queries");
	const sourcePath = join(sourceQueries, "*.pq");
	const project = {
		paths: { dir: root },
		manifest: {
			src: [{ name: "PowerQueries", path: sourcePath }],
			srcProperties: { folder: "src/EmailManager.xlsm", subfolders: { Modules: "Modules" } }
		}
	} as unknown as Project;

	try {
		await mkdir(stagedQueries, { recursive: true });
		await writeFile(join(stagedQueries, "Query.pq"), "query", "utf8");
		jest.clearAllMocks();

		await exportPowerQueries(project, staging);

		expect(await readdir(sourceQueries)).toEqual(["Query.pq"]);
		expect(project.manifest.src).toEqual([{ name: "PowerQueries", path: sourcePath }]);
		expect(project.manifest.srcProperties?.subfolders).toEqual({ Modules: "Modules" });
		expect(writeManifestFile).not.toHaveBeenCalled();
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});
