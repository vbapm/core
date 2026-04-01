import { formatXml } from "../../utils/xml";
import { UnzipFile } from "../../utils/zip";

const XML = /\.(xml|rels)$/i;
const XML_INDENT = 2;

export default function transformFormatXml(file: UnzipFile): UnzipFile {
	if (!XML.test(file.path)) return file;

	try {
		file.data = Buffer.from(formatXml(file.data, { spaces: XML_INDENT }));
	} catch {
		// Some OOXML parts contain content that xml-js cannot parse (e.g. data
		// query files with trailing bytes or a BOM before the declaration).
		// Leave the file unmodified rather than failing the whole export.
		console.warn(`Warning: Failed to parse ${file.path} as XML. Leaving unmodified.`);

	}
	return file;
}
