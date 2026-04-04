import { formatXmlBuffer } from "../../utils/xml";
import { UnzipFile } from "../../utils/zip";

const XML = /\.(xml|rels)$/i;
const XML_INDENT = 2;

export default function transformFormatXml(file: UnzipFile): UnzipFile {
	if (!XML.test(file.path)) return file;

	try {
		file.data = formatXmlBuffer(file.data, { spaces: XML_INDENT }, file.path);
	} catch {
		console.warn(`Warning: Failed to parse ${file.path} as XML. Leaving unmodified.`);
	}
	return file;
}
