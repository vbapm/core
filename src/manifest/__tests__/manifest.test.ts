import { formatManifest, loadManifest, parseManifest } from "../";
import { dev, dir as FIXTURES, invalidManifest, standard } from "../../../tests/__fixtures__";
import { normalizeManifest } from "../../__helpers__/manifest";
import { parse as parseToml } from "../../utils/toml";
import dedent from "@timhall/dedent";

const BASE_MANIFEST: {
	package: {
		name: string;
		version: string;
		authors: string[];
		target?: string | object;
	};
} = {
	package: { name: "package-name", version: "1.0.0", authors: ["Tim Hall"] }
};

test("loads valid package metadata", () => {
	expect(normalizeManifest(parseManifest(BASE_MANIFEST, FIXTURES))).toMatchSnapshot();
});

test("throws for invalid package metadata", () => {
	expect(() => parseManifest({}, FIXTURES)).toThrow();
	expect(() => parseManifest({ package: {} }, FIXTURES)).toThrow();
	expect(() => parseManifest({ package: { name: "package-name" } }, FIXTURES)).toThrow();
	expect(() => parseManifest({ package: { version: "1.0.0" } }, FIXTURES)).toThrow();
	expect(() => parseManifest({ package: { authors: ["Tim Hall"] } }, FIXTURES)).toThrow();
});

test("preserves custom package metadata like license", () => {
	const value = {
		package: {
			...BASE_MANIFEST.package,
			license: "UNLICENSED"
		}
	};

	const manifest = parseManifest(value, FIXTURES);
	expect(manifest.metadata.license).toBe("UNLICENSED");

	const formatted: any = formatManifest(manifest, FIXTURES);
	expect(formatted.package.license).toBe("UNLICENSED");
});

test("loads valid sources", () => {
	const value = {
		...BASE_MANIFEST,
		src: {
			A: "src/a.bas",
			B: { path: "src/b.cls" },
			C: { path: "src/c.frm", optional: true }
		}
	};

	expect(normalizeManifest(parseManifest(value, FIXTURES))).toMatchSnapshot();
});

test("loads src-encoding from [project]", () => {
	const value = {
		...BASE_MANIFEST,
		package: {
			...BASE_MANIFEST.package,
			"src-encoding": "cp1252"
		},
		src: { A: "src/a.bas" }
	};

	const manifest = parseManifest(value, FIXTURES);
	expect(manifest.srcEncoding).toBe("cp1252");
});

test("loads per-source encoding", () => {
	const value = {
		...BASE_MANIFEST,
		src: {
			A: { path: "src/a.bas", encoding: "cp932" }
		}
	};

	const manifest = parseManifest(value, FIXTURES);
	expect(manifest.src[0].encoding).toBe("cp932");
});

test("loads valid [dev-src]", () => {
	const value = {
		...BASE_MANIFEST,
		["dev-src"]: {
			A: "src/a.bas",
			B: { path: "src/b.cls" },
			C: { path: "src/c.frm", optional: true }
		}
	};

	expect(normalizeManifest(parseManifest(value, FIXTURES))).toMatchSnapshot();
});

test("throws for invalid sources", () => {
	const value = {
		...BASE_MANIFEST,
		src: {
			missing_path: { optional: true }
		}
	};

	expect(() => parseManifest(value, FIXTURES)).toThrow();
});

test("loads valid dependencies", () => {
	const value = {
		...BASE_MANIFEST,
		dependencies: {
			a: "^1.0.0",
			b: {
				version: "^2.0.0",
				optional: true,
				features: ["a", "b"],
				"default-features": false
			},
			c: { path: "packages/c" },
			d: { git: "https://github.com/VBA-tools/VBA-Web" },
			e: { git: "https://github.com/VBA-tools/VBA-Web", branch: "next" },
			f: { git: "https://github.com/VBA-tools/VBA-Web", tag: "v1.0.0" },
			g: { git: "https://github.com/VBA-tools/VBA-Web", rev: "a1b2c3d4" }
		}
	};

	expect(normalizeManifest(parseManifest(value, FIXTURES))).toMatchSnapshot();
});

test("loads valid [dev-dependencies]", () => {
	const value = {
		...BASE_MANIFEST,
		"dev-dependencies": {
			a: "^1.0.0",
			b: { version: "^2.0.0" },
			c: { path: "packages/d" },
			d: { git: "https://github.com/VBA-tools/VBA-Web" },
			e: { git: "https://github.com/VBA-tools/VBA-Web", branch: "next" },
			f: { git: "https://github.com/VBA-tools/VBA-Web", tag: "v1.0.0" },
			g: { git: "https://github.com/VBA-tools/VBA-Web", rev: "a1b2c3d4" }
		}
	};

	expect(normalizeManifest(parseManifest(value, FIXTURES))).toMatchSnapshot();
});

test("throws for invalid dependencies", () => {
	const value: any = { ...BASE_MANIFEST, dependencies: { a: {} } };
	expect(() => parseManifest(value, FIXTURES)).toThrow();
});

test("loads valid references", () => {
	const value = {
		...BASE_MANIFEST,
		references: {
			a: {
				version: "1.0",
				guid: "{420B2830-E718-11CF-893D-00A0C9054228}",
				optional: true
			}
		}
	};

	expect(normalizeManifest(parseManifest(value, FIXTURES))).toMatchSnapshot();
});

test("loads valid [dev-references]", () => {
	const value = {
		...BASE_MANIFEST,
		["dev-references"]: {
			a: {
				version: "1.0",
				guid: "{420B2830-E718-11CF-893D-00A0C9054228}",
				optional: true
			}
		}
	};

	expect(normalizeManifest(parseManifest(value, FIXTURES))).toMatchSnapshot();
});

test("throws for invalid references", () => {
	let value: any = { ...BASE_MANIFEST, references: { a: {} } };
	expect(() => parseManifest(value, FIXTURES)).toThrow();

	value = { ...BASE_MANIFEST, references: { b: { version: "1.0.0" } } };
	expect(() => parseManifest(value, FIXTURES)).toThrow();

	value = { ...BASE_MANIFEST, references: { c: { version: "1.0" } } };
	expect(() => parseManifest(value, FIXTURES)).toThrow();

	value = {
		...BASE_MANIFEST,
		references: {
			d: { version: "1.0", guid: "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa" }
		}
	};
	expect(() => parseManifest(value, FIXTURES)).toThrow();
});

test("loads valid target", () => {
	const value = {
		...BASE_MANIFEST
	};
	value.package.target = "xlsm";

	expect(normalizeManifest(parseManifest(value, FIXTURES))).toMatchSnapshot();

	value.package.target = { type: "xlam", name: "addin", path: "targets/xlam" };

	expect(normalizeManifest(parseManifest(value, FIXTURES))).toMatchSnapshot();
});

test("loads target with encoding", () => {
	const value = {
		...BASE_MANIFEST,
		package: {
			...BASE_MANIFEST.package,
			target: { type: "xlsm", encoding: "cp932" }
		}
	};

	const manifest = parseManifest(value, FIXTURES);
	expect(manifest.target?.encoding).toBe("cp932");
});

test("loads and parses manifest", async () => {
	const manifest = await loadManifest(standard);
	expect(normalizeManifest(manifest, standard)).toMatchSnapshot();
});

test("throws for invalid syntax", async () => {
	expect.assertions(1);

	try {
		await loadManifest(invalidManifest);
	} catch (err: any) {
		expect((err?.message || "").replace(FIXTURES, "fixtures")).toMatchSnapshot();
	}
});

test("should format manifest for export", async () => {
	const manifest = await loadManifest(dev);
	const converted = formatManifest(manifest, FIXTURES);

	expect(converted).toMatchSnapshot();
});

describe("section key validation", () => {
	test("rejects snake_case build_dir with suggestion", () => {
		const value = {
			...BASE_MANIFEST,
			package: {
				...BASE_MANIFEST.package,
				build_dir: "."
			}
		};

		expect(() => parseManifest(value, FIXTURES)).toThrow(/Did you mean "build-dir"/);
	});

	test("rejects snake_case src_encoding with suggestion", () => {
		const value = {
			...BASE_MANIFEST,
			package: {
				...BASE_MANIFEST.package,
				src_encoding: "cp1252"
			}
		};

		expect(() => parseManifest(value, FIXTURES)).toThrow(/Did you mean "src-encoding"/);
	});

	test("allows arbitrary metadata keys like license", () => {
		const value = {
			...BASE_MANIFEST,
			package: {
				...BASE_MANIFEST.package,
				license: "MIT",
				description: "A test package"
			}
		};

		// Should not throw — arbitrary metadata is not a misspelled known key
		const manifest = parseManifest(value, FIXTURES);
		expect(manifest.metadata.license).toBe("MIT");
		expect(manifest.metadata.description).toBe("A test package");
	});

	test("rejects snake_case build_dir in [project] section", () => {
		const value = {
			project: {
				name: "test",
				target: "xlsm",
				build_dir: "."
			}
		};

		expect(() => parseManifest(value, FIXTURES)).toThrow(/Did you mean "build-dir"/);
	});
	test("parses wildcard entries as regular source paths", async () => {
		const toml = dedent`
			[project]
			name = "wildcard-test"
			target = "xlsm"

			[src]
			MyStuff = "src/**/*.cls"
			Validation = "src/Validation.bas"
		`;

		const value = await parseToml(toml);
		const manifest = parseManifest(value, FIXTURES);

		// Wildcard stored as literal path, not expanded at parse time
		expect(manifest.src).toHaveLength(2);
		expect(manifest.src[0].name).toBe("MyStuff");
		expect(manifest.src[0].path).toContain("src/**/*.cls");
	});

	test("parses [src-properties] with sort and subfolders", async () => {
		const toml = dedent`
			[project]
			name = "with-props"
			target = "xlsm"

			[src-properties]
			sort.by-types = true
			sort.alphabetical = true

			[src-properties.subfolders]
			Modules = "Modules"
			Classes = "Class Modules"

			[src]
			ModuleA = "src/ModuleA.bas"
		`;

		const value = await parseToml(toml);
		const manifest = parseManifest(value, FIXTURES);

		expect(manifest.srcProperties?.sort?.["by-types"]).toBe(true);
		expect(manifest.srcProperties?.sort?.alphabetical).toBe(true);
		expect(manifest.srcProperties?.subfolders?.Modules).toBe("Modules");
		expect(manifest.srcProperties?.subfolders?.Classes).toBe("Class Modules");
	});
});
