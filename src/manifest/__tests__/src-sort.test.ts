import { detectSrcStructure, parseSrcProperties, resolveSrcFolder } from "../src-sort";
import { Source } from "../source";

function src(name: string, path: string): Source {
	return { name, path };
}

// ---------------------------------------------------------------------------
// detectSrcStructure
// ---------------------------------------------------------------------------

describe("detectSrcStructure", () => {
	test("empty src is unstructured", () => {
		const result = detectSrcStructure([]);
		expect(result).toMatchObject({
			sortedByTypes: false,
			sortedAlphabetically: false,
			sortedByTypeThenAlphabetically: false,
			unstructured: true
		});
	});

	test("sorted by types (bas → frm → cls)", () => {
		const sources: Source[] = [
			src("ModuleA", "src/ModuleA.bas"),
			src("ModuleB", "src/ModuleB.bas"),
			src("FormA", "src/FormA.frm"),
			src("ClassA", "src/ClassA.cls"),
			src("ClassB", "src/ClassB.cls")
		];
		const result = detectSrcStructure(sources);
		expect(result.sortedByTypes).toBe(true);
		expect(result.sortedAlphabetically).toBe(false);
	});

	test("sorted alphabetically (global)", () => {
		const sources: Source[] = [
			src("Alpha", "src/Alpha.bas"),
			src("Beta", "src/Beta.cls"),
			src("Delta", "src/Delta.frm"),
			src("Gamma", "src/Gamma.bas"),
			src("Zeta", "src/Zeta.cls")
		];
		const result = detectSrcStructure(sources);
		expect(result.sortedByTypes).toBe(false);
		expect(result.sortedAlphabetically).toBe(true);
	});

	test("sorted by type then alphabetically", () => {
		const sources: Source[] = [
			src("AlphaModule", "src/AlphaModule.bas"),
			src("BetaModule", "src/BetaModule.bas"),
			src("AlphaForm", "src/AlphaForm.frm"),
			src("BetaForm", "src/BetaForm.frm"),
			src("AlphaClass", "src/AlphaClass.cls"),
			src("BetaClass", "src/BetaClass.cls")
		];
		const result = detectSrcStructure(sources);
		expect(result.sortedByTypes).toBe(true);
		// Not globally alphabetical (Form < Module at boundary), but
		// sortedByTypeThenAlphabetically checks within-each-segment
		expect(result.sortedAlphabetically).toBe(false);
		expect(result.sortedByTypeThenAlphabetically).toBe(true);
	});

	test("unstructured when mix of types is scattered", () => {
		const sources: Source[] = [
			src("ModuleA", "src/ModuleA.bas"),
			src("ClassA", "src/ClassA.cls"),
			src("ModuleB", "src/ModuleB.bas"),
			src("FormA", "src/FormA.frm"),
			src("ClassB", "src/ClassB.cls")
		];
		const result = detectSrcStructure(sources);
		expect(result.sortedByTypes).toBe(false);
		expect(result.unstructured).toBe(true);
	});

	test("too many segments breaks sortedByTypes", () => {
		// 5 segments (cls, bas, frm, cls, bas) > 4 → not sortedByTypes
		const sources: Source[] = [
			src("A", "src/A.cls"),
			src("B", "src/B.bas"),
			src("C", "src/C.frm"),
			src("D", "src/D.cls"),
			src("E", "src/E.bas")
		];
		const result = detectSrcStructure(sources);
		expect(result.sortedByTypes).toBe(false);
	});
});

// ---------------------------------------------------------------------------
// parseSrcProperties
// ---------------------------------------------------------------------------

describe("parseSrcProperties", () => {
	test("returns undefined for absent section", () => {
		expect(parseSrcProperties(undefined)).toBeUndefined();
		expect(parseSrcProperties(null)).toBeUndefined();
		expect(parseSrcProperties({})).toBeUndefined();
	});

	test("parses sort options", () => {
		const result = parseSrcProperties({
			sort: { "by-types": true, alphabetical: true }
		});
		expect(result).toEqual({
			sort: { "by-types": true, alphabetical: true }
		});
	});

	test("parses partial sort options", () => {
		const result = parseSrcProperties({
			sort: { "by-types": true }
		});
		expect(result).toEqual({
			sort: { "by-types": true }
		});
	});

	test("ignores unknown keys", () => {
		const result = parseSrcProperties({
			unknown: "value"
		});
		expect(result).toBeUndefined();
	});

	test("returns undefined when all values are invalid types", () => {
		const result = parseSrcProperties({
			sort: "not-an-object"
		});
		expect(result).toBeUndefined();
	});

	test("parses folder", () => {
		const result = parseSrcProperties({ folder: "src/MyWorkbook" });
		expect(result).toEqual({ folder: "src/MyWorkbook" });
	});

	test("parses folder alongside other options", () => {
		const result = parseSrcProperties({
			folder: "src/MyWorkbook",
			sort: { "by-types": true }
		});
		expect(result).toEqual({
			folder: "src/MyWorkbook",
			sort: { "by-types": true }
		});
	});
});

// ---------------------------------------------------------------------------
// resolveSrcFolder
// ---------------------------------------------------------------------------

describe("resolveSrcFolder", () => {
	test('returns "src" when srcProperties is undefined', () => {
		expect(resolveSrcFolder(undefined)).toBe("src");
	});

	test('returns "src" when folder is not set', () => {
		expect(resolveSrcFolder({})).toBe("src");
	});

	test("returns the explicit folder value", () => {
		expect(resolveSrcFolder({ folder: "lib" })).toBe("lib");
	});
});
