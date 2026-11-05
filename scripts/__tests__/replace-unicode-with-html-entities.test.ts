import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
	replaceFileContents,
	replaceWithNamedHtmlEntities
} from "../replace-unicode-with-html-entities.js";
import { createNamedHtmlEntityInvalidCharHandler } from "../html-entities.js";

describe("replaceWithNamedHtmlEntities", () => {
	test("uses the semicolon-terminated named reference for ö", () => {
		expect(replaceWithNamedHtmlEntities(String.fromCodePoint(0xf6))).toEqual({
			text: "&ouml;",
			replacements: 1,
			unconverted: 0
		});
	});

	test("matches multi-codepoint entities and leaves existing ASCII references alone", () => {
		const text = `&ouml; ${String.fromCodePoint(0x2242, 0x338)} A`;

		expect(replaceWithNamedHtmlEntities(text)).toEqual({
			text: "&ouml; &nesim; A",
			replacements: 1,
			unconverted: 0
		});
	});

	test("leaves Unicode characters without a named reference unchanged", () => {
		const text = String.fromCodePoint(0x1f600);

		expect(replaceWithNamedHtmlEntities(text)).toEqual({
			text,
			replacements: 0,
			unconverted: 1
		});
	});
});

test("reports named-sequence and numeric substitutions at their original character indices", () => {
	const entities = {
		"&ouml;": { codepoints: [246], characters: "ö" },
		"&nesim;": { codepoints: [8770, 824], characters: String.fromCodePoint(0x2242, 0x338) }
	};
	const text = ` \t' Aö${entities["&nesim;"].characters} ${String.fromCodePoint(0xac00)}`;
	const substitutions: Array<[string, string, number]> = [];
	const handler = createNamedHtmlEntityInvalidCharHandler(entities, text, (...event) => {
		substitutions.push(event);
	});

	expect(handler("ö", 5)).toBe("&ouml;");
	expect(handler("≂", 6)).toBe("&nesim;");
	expect(handler("̸", 7)).toBe("");
	expect(handler("가", 9)).toBe("&#xac00;");
	expect(substitutions).toEqual([
		["ö", "&ouml;", 5],
		[entities["&nesim;"].characters, "&nesim;", 6],
		["가", "&#xac00;", 9]
	]);
});

test("refuses to replace characters outside apostrophe comment lines", () => {
	const entities = { "&ouml;": { codepoints: [246], characters: "ö" } };
	const text = `Public Const Name = "Dörring"`;
	const handler = createNamedHtmlEntityInvalidCharHandler(entities, text);

	expect(() => handler("ö", text.indexOf("ö"))).toThrow(
		/it is not inside an apostrophe comment, and replacing it could alter code behavior/
	);
});

describe("replaceFileContents", () => {
	let directory: string;

	afterEach(() => {
		rmSync(directory, { recursive: true, force: true });
	});

	test("preserves a UTF-8 BOM and CRLF line endings", () => {
		directory = mkdtempSync(join(tmpdir(), "html-entities-"));
		const filePath = join(directory, "source.txt");
		const original = Buffer.concat([
			Buffer.from([0xef, 0xbb, 0xbf]),
			Buffer.from(`Name: ${String.fromCodePoint(0xf6)}\r\n`, "utf8")
		]);
		writeFileSync(filePath, original);

		expect(replaceFileContents(filePath)).toEqual({ replacements: 1, unconverted: 0 });
		expect(readFileSync(filePath)).toEqual(
			Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from("Name: &ouml;\r\n", "utf8")])
		);
	});

	test("rejects files that are not valid UTF-8", () => {
		directory = mkdtempSync(join(tmpdir(), "html-entities-"));
		const filePath = join(directory, "invalid.txt");
		writeFileSync(filePath, Buffer.from([0xff, 0xfe, 0xfd]));

		expect(() => replaceFileContents(filePath)).toThrow("File is not valid UTF-8");
	});
});
