import { formatXml } from "../../utils/xml";
import { UnzipFile } from "../../utils/zip";

const XML = /\.(xml|rels)$/i;
const XML_INDENT = 2;

export default function transformFormatXml(file: UnzipFile): UnzipFile {
	if (!XML.test(file.path)) return file;

	file.data = Buffer.from(formatXml(file.data, { spaces: XML_INDENT }));
	return file;
}
