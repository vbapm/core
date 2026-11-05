import { mkdtemp, mkdir, readFile, readdir, rm, writeFile } from "fs/promises";
import { tmpdir } from "os";
import { join } from "path";
import { resolveSourceFiles } from "../../build/apply-extract";
import { writeFile as writeManifestFile } from "../../utils/fs";
import { Project } from "../../project";
import { loadPowerQueries } from "../../build/load-power-queries";
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
		manifest: {
			type: "project",
			name: "power-query-test",
			version: "0.0.0",
			metadata: { authors: [] },
			src: [],
			devSrc: [],
			dependencies: [],
			references: [],
			devDependencies: [],
			devReferences: [],
			srcProperties: undefined
		}
	} as unknown as Project;

	try {
		await mkdir(stagedQueries, { recursive: true });
		await writeFile(join(stagedQueries, "Query.pq"), "query", "utf8");

		await exportPowerQueries(project, staging);

		expect(await readdir(fallbackQueries)).toEqual(["Query.pq"]);
		expect(project.manifest.src).toEqual([
			{ name: "Query", path: join(fallbackQueries, "Query.pq") }
		]);
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

test("uses an unused source key when PowerQueries is already a VBA source", async () => {
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-"));
	const staging = join(root, "staging");
	const stagedQueries = join(staging, "pq");
	const existingVbaSource = { name: "PowerQueries", path: join(root, "src", "PowerQueries.bas") };
	const project = {
		paths: { dir: root },
		manifest: {
			type: "project",
			name: "power-query-test",
			version: "0.0.0",
			metadata: { authors: [] },
			src: [existingVbaSource],
			srcProperties: { folder: "src", subfolders: { Modules: "Modules" } },
			dependencies: [],
			references: [],
			devSrc: [],
			devDependencies: [],
			devReferences: []
		}
	} as unknown as Project;

	try {
		await mkdir(stagedQueries, { recursive: true });
		await writeFile(join(stagedQueries, "Query.pq"), "query", "utf8");
		jest.clearAllMocks();

		await exportPowerQueries(project, staging);

		expect(project.manifest.src).toEqual([
			existingVbaSource,
			{ name: "PowerQueries2", path: join(root, "src", "Power Queries", "*.pq") }
		]);
		expect(writeManifestFile).toHaveBeenCalledWith(
			expect.stringContaining("vbaproject.toml"),
			expect.stringContaining('PowerQueries2 = "src/Power Queries/*.pq"')
		);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});

test("writes an explicitly mapped query and preserves unrelated query files", async () => {
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-"));
	const staging = join(root, "staging");
	const stagedQueries = join(staging, "pq");
	const queryDirectory = join(root, "src", "queries");
	const mappedPath = join(queryDirectory, "stored-name.pq");
	const staleMappedPath = join(queryDirectory, "Removed.pq");
	const unrelatedPath = join(queryDirectory, "Unrelated.pq");
	const project = {
		paths: { dir: root },
		manifest: {
			type: "project",
			name: "power-query-test",
			version: "0.0.0",
			metadata: { authors: [] },
			src: [
				{ name: "Query", path: mappedPath },
				{ name: "Removed", path: staleMappedPath }
			],
			devSrc: [],
			dependencies: [],
			references: [],
			devDependencies: [],
			devReferences: [],
			srcProperties: { subfolders: { PowerQueries: "queries" } }
		}
	} as unknown as Project;

	try {
		await mkdir(stagedQueries, { recursive: true });
		await mkdir(queryDirectory, { recursive: true });
		await writeFile(join(stagedQueries, "Query.pq"), "updated query", "utf8");
		await writeFile(staleMappedPath, "remove this mapped query", "utf8");
		await writeFile(unrelatedPath, "keep this query", "utf8");

		await exportPowerQueries(project, staging);

		expect(await readFile(mappedPath, "utf8")).toBe("updated query");
		expect(await readFile(unrelatedPath, "utf8")).toBe("keep this query");
		await expect(readFile(staleMappedPath, "utf8")).rejects.toMatchObject({ code: "ENOENT" });
		expect(project.manifest.src).toEqual([{ name: "Query", path: mappedPath }]);
		expect(writeManifestFile).toHaveBeenCalled();
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});

test("updates dev Power Query sources and preserves dev queries absent from the workbook", async () => {
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-dev-"));
	const staging = join(root, "staging");
	const stagedQueries = join(staging, "pq");
	const devQueries = join(root, "dev-src", "pq");
	const project = {
		paths: { dir: root },
		manifest: {
			src: [],
			devSrc: [{ name: "DevQueries", path: join(devQueries, "*.pq") }],
			srcProperties: undefined
		}
	} as unknown as Project;

	try {
		await mkdir(stagedQueries, { recursive: true });
		await mkdir(devQueries, { recursive: true });
		await writeFile(join(stagedQueries, "DevQuery.pq"), "updated dev query", "utf8");
		await writeFile(join(devQueries, "DevQuery.pq"), "old dev query", "utf8");
		await writeFile(join(devQueries, "DevOnly.pq"), "keep dev query", "utf8");

		await exportPowerQueries(project, staging);

		expect(await readFile(join(devQueries, "DevQuery.pq"), "utf8")).toBe("updated dev query");
		expect(await readFile(join(devQueries, "DevOnly.pq"), "utf8")).toBe("keep dev query");
		expect(project.manifest.src).toEqual([]);
		expect(project.manifest.devSrc).toEqual([
			{ name: "DevQueries", path: join(devQueries, "*.pq") }
		]);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});

test("preserves the original query name when the staged filename is sanitized", async () => {
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-name-"));
	const staging = join(root, "staging");
	const stagedQueries = join(staging, "pq");
	const queryFolder = join(root, "src", "pq");
	const queryPath = join(queryFolder, "_NUL.pq");
	const project = {
		paths: { dir: root },
		manifest: {
			type: "project",
			name: "power-query-test",
			version: "0.0.0",
			metadata: { authors: [] },
			src: [{ name: "PowerQueries", path: join(queryFolder, "*.pq") }],
			devSrc: [],
			dependencies: [],
			references: [],
			devDependencies: [],
			devReferences: [],
			srcProperties: undefined
		}
	} as unknown as Project;

	try {
		await mkdir(stagedQueries, { recursive: true });
		await mkdir(queryFolder, { recursive: true });
		await writeFile(join(stagedQueries, "_NUL.pq"), "updated reserved-name query", "utf8");
		await writeFile(
			join(staging, "project.json"),
			JSON.stringify({
				powerQueries: [{ name: "NUL", file: "_NUL.pq" }]
			})
		);
		await writeFile(queryPath, "old query", "utf8");

		await exportPowerQueries(project, staging);

		expect(await readFile(queryPath, "utf8")).toBe("updated reserved-name query");
		await expect(loadPowerQueries(project)).resolves.toEqual([
			{ name: "NUL", formula: "updated reserved-name query" }
		]);
		expect(project.manifest.src).toEqual([
			{ name: "PowerQueries", path: join(queryFolder, "*.pq") },
			{ name: "NUL", path: queryPath }
		]);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});

test("registers a query when an existing wildcard does not match its filename", async () => {
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-pattern-"));
	const staging = join(root, "staging");
	const stagedQueries = join(staging, "pq");
	const queryDirectory = join(root, "src", "pq");
	const project = {
		paths: { dir: root },
		manifest: {
			type: "project",
			name: "power-query-test",
			version: "0.0.0",
			metadata: { authors: [] },
			src: [{ name: "SalesQueries", path: join(queryDirectory, "Sales*.pq") }],
			devSrc: [],
			dependencies: [],
			references: [],
			devDependencies: [],
			devReferences: [],
			srcProperties: undefined
		}
	} as unknown as Project;

	try {
		await mkdir(stagedQueries, { recursive: true });
		await writeFile(join(stagedQueries, "Inventory.pq"), "inventory query", "utf8");

		await exportPowerQueries(project, staging);

		expect(project.manifest.src).toEqual([
			{ name: "SalesQueries", path: join(queryDirectory, "Sales*.pq") },
			{ name: "Inventory", path: join(queryDirectory, "Inventory.pq") }
		]);
		await expect(loadPowerQueries(project)).resolves.toEqual([
			{ name: "Inventory", formula: "inventory query" }
		]);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});

test("updates existing nested matches for a recursive Power Query source", async () => {
	const root = await mkdtemp(join(tmpdir(), "vbapm-power-query-"));
	const staging = join(root, "staging");
	const stagedQueries = join(staging, "pq");
	const queryRoot = join(root, "src", "Power Queries");
	const nestedDirectory = join(queryRoot, "Customer", "Reports");
	const existingPath = join(nestedDirectory, "Existing.pq");
	const stalePath = join(nestedDirectory, "Stale.pq");
	const sourcePath = join(queryRoot, "**", "*.pq");
	const project = {
		paths: { dir: root },
		manifest: {
			src: [{ name: "PowerQueries", path: sourcePath }],
			srcProperties: { subfolders: { PowerQueries: "Power Queries" } }
		}
	} as unknown as Project;

	try {
		await mkdir(stagedQueries, { recursive: true });
		await mkdir(nestedDirectory, { recursive: true });
		await writeFile(join(stagedQueries, "Existing.pq"), "updated nested query", "utf8");
		await writeFile(join(stagedQueries, "New.pq"), "new query", "utf8");
		await writeFile(existingPath, "old query", "utf8");
		await writeFile(stalePath, "stale query", "utf8");

		await exportPowerQueries(project, staging);

		expect(await readFile(existingPath, "utf8")).toBe("updated nested query");
		expect(await readFile(join(queryRoot, "New.pq"), "utf8")).toBe("new query");
		await expect(readFile(stalePath, "utf8")).rejects.toMatchObject({ code: "ENOENT" });
		expect(project.manifest.src).toEqual([{ name: "PowerQueries", path: sourcePath }]);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});
