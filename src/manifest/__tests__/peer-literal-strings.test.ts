import { toPeerLiteralStrings } from "../index";

describe("toPeerLiteralStrings", () => {
	test("converts peer path basic string to literal string", () => {
		const input = `[references]\nAddinToolbox = { peer = true, path = "C:\\\\Users\\\\leduc\\\\AddinToolbox.xlam" }`;

		expect(toPeerLiteralStrings(input)).toBe(
			`[references]\nAddinToolbox = { peer = true, path = 'C:\\Users\\leduc\\AddinToolbox.xlam' }`
		);
	});

	test("leaves COM references untouched", () => {
		const input = `[references]\nScripting = { guid = "{420B2830-E718-11CF-893D-00A0C9054228}", version = "1.0" }`;

		expect(toPeerLiteralStrings(input)).toBe(input);
	});

	test("keeps basic string when value contains a single quote", () => {
		const input = `AddinToolbox = { peer = true, path = "C:\\\\Users\\\\O'Brien\\\\AddinToolbox.xlam" }`;

		expect(toPeerLiteralStrings(input)).toBe(input);
	});

	test("keeps basic string when value contains non-path escapes", () => {
		const input = `AddinToolbox = { peer = true, path = "C:\\\\Users\\\\x\\ty.xlam" }`;

		expect(toPeerLiteralStrings(input)).toBe(input);
	});

	test("handles relative peer paths without backslashes", () => {
		const input = `AddinToolbox = { peer = true, path = "../AddinToolbox/build/AddinToolbox.xlam" }`;

		expect(toPeerLiteralStrings(input)).toBe(
			`AddinToolbox = { peer = true, path = '../AddinToolbox/build/AddinToolbox.xlam' }`
		);
	});

	test("converts multiple peer paths on separate lines", () => {
		const input = [
			`A = { peer = true, path = "C:\\\\a.xlam" }`,
			`B = { peer = true, path = "C:\\\\b.xlam" }`
		].join("\n");

		const expected = [`A = { peer = true, path = 'C:\\a.xlam' }`, `B = { peer = true, path = 'C:\\b.xlam' }`].join(
			"\n"
		);

		expect(toPeerLiteralStrings(input)).toBe(expected);
	});
});
