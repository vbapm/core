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
});
