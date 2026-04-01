import transformFormatXml from "../format-xml";
import { UnzipFile } from "../../../utils/zip";

function toFile(path: string, data: string): UnzipFile {
	return {
		path,
		data: Buffer.from(data),
		mode: 0,
		mtime: "",
		type: "file"
	};
}

function toFileBuffer(path: string, data: Buffer): UnzipFile {
	return {
		path,
		data,
		mode: 0,
		mtime: "",
		type: "file"
	};
}

describe("transformFormatXml", () => {
	test("should format xml with stable indentation", () => {
		const file = toFile("xl/workbook.xml", "<root><child>value</child></root>");

		const transformed = transformFormatXml(file);
		const result = transformed.data.toString("utf8");

		expect(result).toContain("\n");
		expect(result).toContain("  <child>");
	});

	test("should not change non-xml files", () => {
		const original = "module text";
		const file = toFile("xl/vbaProject.bin", original);

		const transformed = transformFormatXml(file);

		expect(transformed.data.toString("utf8")).toBe(original);
	});

	test("should leave unparseable xml files unmodified", () => {
		// Some OOXML parts (e.g. data query files) contain content that xml-js
		// cannot parse. The transform should not throw and should return the
		// original content unchanged.
		const original = "not valid xml \x00\x01\x02";
		const file = toFile("xl/queryTables/queryTable1.xml", original);

		expect(() => transformFormatXml(file)).not.toThrow();
		expect(
			transformFormatXml(toFile("xl/queryTables/queryTable1.xml", original)).data.toString("utf8")
		).toBe(original);
	});

	test("should parse utf-16le xml input", () => {
		const xml = '<?xml version="1.0" encoding="utf-16"?><root><child>value</child></root>';
		const utf16le = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(xml, "utf16le")]);
		const file = toFileBuffer("xl/workbook.xml", utf16le);

		expect(() => transformFormatXml(file)).not.toThrow();
		expect(transformFormatXml(file).data.toString("utf8")).toContain("<child>value</child>");
	});
});
