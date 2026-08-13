import {
	formatReferences,
	parseReference,
	parseReferences,
	relativizePeerPath,
	resolvePeerReferencePaths,
	Reference
} from "../reference";

describe("parseReference", () => {
	test("parses peer reference without guid/version", () => {
		const ref = parseReference("AddinToolbox", { peer: true });

		expect(ref).toEqual({
			name: "AddinToolbox",
			guid: "",
			major: 0,
			minor: 0,
			peer: true
		});
	});

	test("parses peer reference with path", () => {
		const ref = parseReference("AddinToolbox", { peer: true, path: "../AddinToolbox/build/AddinToolbox.xlam" });

		expect(ref.path).toBe("../AddinToolbox/build/AddinToolbox.xlam");
		expect(ref.peer).toBe(true);
	});

	test("still parses COM reference", () => {
		const ref = parseReference("Scripting", {
			version: "1.0",
			guid: "{420B2830-E718-11CF-893D-00A0C9054228}"
		});

		expect(ref).toEqual({
			name: "Scripting",
			guid: "{420B2830-E718-11CF-893D-00A0C9054228}",
			major: 1,
			minor: 0
		});
	});
});

describe("formatReferences", () => {
	test("formats peer reference as peer = true", () => {
		const refs: Reference[] = [
			{ name: "AddinToolbox", guid: "", major: 0, minor: 0, peer: true }
		];

		expect(formatReferences(refs)).toEqual({
			AddinToolbox: { peer: true }
		});
	});

	test("formats peer reference with path", () => {
		const refs: Reference[] = [
			{
				name: "AddinToolbox",
				guid: "",
				major: 0,
				minor: 0,
				peer: true,
				path: "../AddinToolbox/build/AddinToolbox.xlam"
			}
		];

		expect(formatReferences(refs)).toEqual({
			AddinToolbox: { peer: true, path: "../AddinToolbox/build/AddinToolbox.xlam" }
		});
	});

	test("keeps COM references unchanged", () => {
		const refs: Reference[] = [
			{
				name: "Scripting",
				guid: "{420B2830-E718-11CF-893D-00A0C9054228}",
				major: 1,
				minor: 0
			}
		];

		expect(formatReferences(refs)).toEqual({
			Scripting: { version: "1.0", guid: "{420B2830-E718-11CF-893D-00A0C9054228}" }
		});
	});

	test("round-trips peer reference through parse and format", () => {
		const toml = { AddinToolbox: { peer: true, path: "../AddinToolbox.xlam" } };
		const parsed = parseReferences(toml);
		expect(formatReferences(parsed)).toEqual(toml);
	});
});

describe("relativizePeerPath", () => {
	const projectDir = "C:/Users/alice/projects/my-app";

	test("keeps relative path when peer is inside project folder", () => {
		expect(relativizePeerPath(projectDir, `${projectDir}/build/AddinToolbox.xlam`)).toBe(
			"build/AddinToolbox.xlam"
		);
	});

	test("keeps relative path when peer is in a sibling folder", () => {
		expect(relativizePeerPath(projectDir, "C:/Users/alice/projects/AddinToolbox/build/AddinToolbox.xlam")).toBe(
			"../AddinToolbox/build/AddinToolbox.xlam"
		);
	});

	test("keeps absolute path when peer is far away", () => {
		const peer = "C:/Users/alice/other-location/AddinToolbox.xlam";
		expect(relativizePeerPath(projectDir, peer)).toBe(peer);
	});

	test("passes through already-relative paths", () => {
		expect(relativizePeerPath(projectDir, "../AddinToolbox.xlam")).toBe("../AddinToolbox.xlam");
	});
});

describe("resolvePeerReferencePaths", () => {
	const projectDir = "C:/Users/alice/projects/my-app";

	test("resolves relative peer path against project folder", () => {
		const refs: Reference[] = [
			{
				name: "AddinToolbox",
				guid: "",
				major: 0,
				minor: 0,
				peer: true,
				path: "../AddinToolbox/build/AddinToolbox.xlam"
			}
		];

		expect(resolvePeerReferencePaths(refs, projectDir)).toEqual([
			{
				name: "AddinToolbox",
				guid: "",
				major: 0,
				minor: 0,
				peer: true,
				path: "C:/Users/alice/projects/AddinToolbox/build/AddinToolbox.xlam"
			}
		]);
	});

	test("keeps absolute peer path unchanged", () => {
		const refs: Reference[] = [
			{
				name: "AddinToolbox",
				guid: "",
				major: 0,
				minor: 0,
				peer: true,
				path: "C:/Users/alice/other/AddinToolbox.xlam"
			}
		];

		expect(resolvePeerReferencePaths(refs, projectDir)).toEqual(refs);
	});

	test("passes through COM references unchanged", () => {
		const refs: Reference[] = [
			{
				name: "Scripting",
				guid: "{420B2830-E718-11CF-893D-00A0C9054228}",
				major: 1,
				minor: 0
			}
		];

		expect(resolvePeerReferencePaths(refs, projectDir)).toEqual(refs);
	});

	test("passes through peer reference without path", () => {
		const refs: Reference[] = [{ name: "AddinToolbox", guid: "", major: 0, minor: 0, peer: true }];

		expect(resolvePeerReferencePaths(refs, projectDir)).toEqual(refs);
	});
});
