import { xml2js, js2xml, Element } from "xml-js";

export type Xml = any;

function decodeXmlBuffer(xml: Buffer): string {
	if (xml.length >= 2) {
		if (xml[0] === 0xff && xml[1] === 0xfe) {
			return xml.toString("utf16le");
		}

		if (xml[0] === 0xfe && xml[1] === 0xff) {
			// Node does not support utf16be directly.
			const littleEndian = Buffer.allocUnsafe(xml.length);
			for (let i = 0; i < xml.length - 1; i += 2) {
				littleEndian[i] = xml[i + 1];
				littleEndian[i + 1] = xml[i];
			}

			if (xml.length % 2 === 1) {
				littleEndian[xml.length - 1] = xml[xml.length - 1];
			}

			return littleEndian.toString("utf16le");
		}
	}

	let oddNulls = 0;
	let evenNulls = 0;
	for (let i = 0; i < xml.length; i++) {
		if (xml[i] !== 0x00) continue;
		if (i % 2 === 0) {
			evenNulls++;
		} else {
			oddNulls++;
		}
	}

	// UTF-16 XML without BOM often has null bytes at every other position.
	if (oddNulls > 8 && oddNulls > evenNulls * 2) {
		return xml.toString("utf16le");
	}

	if (evenNulls > 8 && evenNulls > oddNulls * 2) {
		const littleEndian = Buffer.allocUnsafe(xml.length);
		for (let i = 0; i < xml.length - 1; i += 2) {
			littleEndian[i] = xml[i + 1];
			littleEndian[i + 1] = xml[i];
		}

		if (xml.length % 2 === 1) {
			littleEndian[xml.length - 1] = xml[xml.length - 1];
		}

		return littleEndian.toString("utf16le");
	}

	return xml.toString("utf8");
}

export function parseXml(xml: string | Buffer): Element {
	if (Buffer.isBuffer(xml)) {
		xml = decodeXmlBuffer(xml);
	}

	return xml2js(xml, { compact: false }) as Element;
}

export interface ConvertOptions {
	// https://github.com/nashwaan/xml-js#options-for-converting-js-object--json--xml
	compact?: boolean;
	spaces?: number | string;
}

export function convertXml(value: Xml, options?: ConvertOptions): string {
	return js2xml(value, options);
}

export function formatXml(xml: string | Buffer, options: ConvertOptions = {}): string {
	return convertXml(parseXml(xml), options);
}

export function findElement(
	elements: Element[] | undefined,
	callback: (element: Element, index: number, elements: Element[]) => boolean
): Element | undefined {
	return elements && elements.find(callback);
}

export function findElementByName(elements: Element[] | undefined, name: string) {
	return findElement(elements, element => element.name === name);
}
