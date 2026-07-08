import { detectSrcStructure, parseSrcProperties, SrcStructure } from "../src-sort";
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
			grouped: false,
			sortedByTypes: false,
			sortedAlphabetically: false,
			sortedByTypeThenAlphabetically: false,
			unstructured: true
		});
	});

	test("detects grouped convention with 4 keys (Objects + Modules + Forms + Classes)", () => {
		const sources: Source[] = [
			src("Objects", "src/Excel Objects/*.cls"),
			src("Modules", "src/**/*.bas"),
			src("Forms", "src/**/*.frm"),
			src("Classes", "src/Class Modules/*.cls")
		];
		const result = detectSrcStructure(sources);
		expect(result.grouped).toBe(true);
		expect(result.groupedPatterns).toEqual({
			Objects: "src/Excel Objects/*.cls",
			Modules: "src/**/*.bas",
			Forms: "src/**/*.frm",
			Classes: "src/Class Modules/*.cls"
		});
	});

	test("detects grouped convention with 3 keys (no Objects)", () => {
		const sources: Source[] = [
			src("Modules", "src/**/*.bas"),
			src("Forms", "src/**/*.frm"),
			src("Classes", "src/**/*.cls")
		];
		const result = detectSrcStructure(sources);
		expect(result.grouped).toBe(true);
		expect(result.groupedPatterns).toEqual({
			Modules: "src/**/*.bas",
			Forms: "src/**/*.frm",
			Classes: "src/**/*.cls"
		});
	});

	test("detects grouped convention with array patterns", () => {
		const sources: Source[] = [
			src("Modules", "src/**/*.bas"),
			src("Forms", "src/Forms/*.frm"),
			src("Classes", "src/Classes/*.cls")
		];
		const result = detectSrcStructure(sources);
		expect(result.grouped).toBe(true);
	});

	test("grouped detection is case-insensitive for key names", () => {
		const sources: Source[] = [
			src("modules", "src/**/*.bas"),
			src("FORMS", "src/**/*.frm"),
			src("Classes", "src/**/*.cls")
		];
		const result = detectSrcStructure(sources);
		expect(result.grouped).toBe(true);
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
		expect(result.grouped).toBe(false);
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

	test("parses grouping flag", () => {
		const result = parseSrcProperties({ grouping: true });
		expect(result).toEqual({ grouping: true });
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
			grouping: true,
			unknown: "value"
		});
		expect(result).toEqual({ grouping: true });
	});

	test("returns undefined when all values are invalid types", () => {
		const result = parseSrcProperties({
			grouping: "not-a-boolean",
			sort: "not-an-object"
		});
		expect(result).toBeUndefined();
	});
});
