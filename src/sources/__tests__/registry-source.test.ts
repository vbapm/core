import { mkdirSync, readdirSync, writeFileSync } from "fs";
import { join } from "../../utils/path";
import { pathExists, tmpFolder } from "../../utils/fs";
import {
	extractSource,
	sanitizePackageName,
	getRemotePackage,
	getLocalPackage,
	getSource
} from "../registry-source";

jest.mock("../../utils/zip", () => ({
	unzip: jest.fn(async (_file: string, dest: string) => {
		// Simulate `decompress` writing the package's legacy manifest.
		mkdirSync(dest, { recursive: true });
		writeFileSync(join(dest, "vba-block.toml"), '[package]\nname = "dictionary"\n');
	})
}));

test("should resolve registry dependency", () => {
	// TODO
});

test("should fetch registry dependency", () => {
	// TODO
});

describe("utils", () => {
	const dictionary = {
		name: "dictionary",
		version: "1.0.0",
		dependencies: [],
		id: "dictionary@1.0.0",
		source: ""
	};

	test("should sanitize package name", () => {
		expect(sanitizePackageName("vba-tools/log")).toEqual("vba-tools--log");
		expect(sanitizePackageName("vba/tools--log")).toEqual("vba--tools--log");
		expect(sanitizePackageName('a/b.c\\d:e*f"g>h<i|j')).toEqual("a--b.c-d-e-f-g-h-i-j");
	});

	test("should get remote package url", () => {
		expect(getRemotePackage("https://packages.vba-blocks.com", dictionary)).toEqual(
			"https://packages.vba-blocks.com/dictionary-v1.0.0.block"
		);
	});

	test("should get local package path", () => {
		expect(getLocalPackage(".vbapm/packages", dictionary)).toEqual(
			".vbapm/packages/dictionary-v1.0.0.block"
		);
	});

	test("should get source path", () => {
		expect(getSource(".vbapm/sources", dictionary)).toEqual(".vbapm/sources/dictionary-v1.0.0");
	});
});

describe("extractSource", () => {
	const dictionary = {
		name: "dictionary",
		version: "1.0.0",
		dependencies: [],
		id: "dictionary@1.0.0",
		source: ""
	};

	test("should extract the same package safely when fetched concurrently", async () => {
		const root = await tmpFolder({ prefix: "registry-source-test-" });
		const sources = join(root, "sources");
		const file = join(root, "dictionary-v1.0.0.block");
		const src = join(sources, "dictionary-v1.0.0");

		const [a, b] = await Promise.all([
			extractSource(file, sources, dictionary),
			extractSource(file, sources, dictionary)
		]);

		expect(a).toBe(src);
		expect(b).toBe(src);

		// The legacy manifest must be normalized exactly once, with no leftovers.
		expect(await pathExists(join(src, "vbaproject.toml"))).toBe(true);
		expect(await pathExists(join(src, "vba-block.toml"))).toBe(false);

		// No temp extraction dirs should be left behind.
		expect(readdirSync(sources).sort()).toEqual(["dictionary-v1.0.0"]);
	});
});
