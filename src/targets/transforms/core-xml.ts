import { parseXml, convertXml, findElementByName } from "../../utils/xml";
import { env } from "../../env";

import { UnzipFile } from "../../utils/zip";

const debug = env.debug("vbapm:target.transforms.core-xml");
const CORE_XML = /docProps[\/,\\]core\.xml/i;

export default function transformCoreXml(file: UnzipFile): UnzipFile {
	if (!CORE_XML.test(file.path)) return file;

	// core.xml is always utf-8 encoded, so we can directly parse the buffer as utf-8 string without needing to detect encoding.
	const xml = parseXml(file.data.toString("utf8"));

	// 1. cp:coreProperties > cp:lastModifiedBy -> Replace with dc:creator
	// 2. cp:coreProperties > dcterms:modified -> Replace with dcterms:created
	const core_properties = findElementByName(xml.elements, "cp:coreProperties");
	if (core_properties) {
		// 1.
		const last_modified_by = findElementByName(core_properties.elements, "cp:lastModifiedBy");
		if (last_modified_by && last_modified_by.elements && last_modified_by.elements[0]) {
			const creator = findElementByName(core_properties.elements, "dc:creator");
			const creatorText =
				creator && creator.elements && creator.elements[0] ? creator.elements[0].text : "";
			last_modified_by.elements[0].text = creatorText;
		}

		// 2.
		const modified = findElementByName(core_properties.elements, "dcterms:modified");
		if (modified && modified.elements && modified.elements[0]) {
			const created = findElementByName(core_properties.elements, "dcterms:created");
			const createdText =
				created && created.elements && created.elements[0] ? created.elements[0].text : "";
			modified.elements[0].text = createdText;
		}
	} else {
		debug("Warning: cp:coreProperties not found, unable to transform core.xml");
	}

	file.data = Buffer.from(convertXml(xml));
	return file;
}
