import { xml2js, js2xml, Element } from "xml-js";

export type Xml = any;

type XmlEncoding = "utf8" | "utf16le" | "utf16be";

interface XmlBufferEncoding {
	encoding: XmlEncoding;
	hasBom: boolean;
}

function swap16Bytes(buffer: Buffer): Buffer {
	const swapped = Buffer.allocUnsafe(buffer.length);
	for (let i = 0; i < buffer.length - 1; i += 2) {
		swapped[i] = buffer[i + 1];
		swapped[i + 1] = buffer[i];
	}

	if (buffer.length % 2 === 1) {
		swapped[buffer.length - 1] = buffer[buffer.length - 1];
	}

	return swapped;
}

function detectXmlBufferEncoding(xml: Buffer): XmlBufferEncoding {
	if (xml.length >= 2) {
		if (xml[0] === 0xff && xml[1] === 0xfe) {
			return { encoding: "utf16le", hasBom: true };
		}

		if (xml[0] === 0xfe && xml[1] === 0xff) {
			return { encoding: "utf16be", hasBom: true };
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
		return { encoding: "utf16le", hasBom: false };
	}

	if (evenNulls > 8 && evenNulls > oddNulls * 2) {
		return { encoding: "utf16be", hasBom: false };
	}

	return { encoding: "utf8", hasBom: false };
}

function decodeXmlBuffer(xml: Buffer): { text: string; encoding: XmlBufferEncoding } {
	const encoding = detectXmlBufferEncoding(xml);

	if (encoding.encoding === "utf16le") {
		return { text: xml.toString("utf16le"), encoding };
	}

	if (encoding.encoding === "utf16be") {
		return { text: swap16Bytes(xml).toString("utf16le"), encoding };
	}

	return { text: xml.toString("utf8"), encoding };
}

function encodeXmlBuffer(xml: string, encoding: XmlBufferEncoding): Buffer {
	if (encoding.encoding === "utf16le") {
		const data = Buffer.from(xml, "utf16le");
		if (!encoding.hasBom) return data;
		return Buffer.concat([Buffer.from([0xff, 0xfe]), data]);
	}

	if (encoding.encoding === "utf16be") {
		const data = swap16Bytes(Buffer.from(xml, "utf16le"));
		if (!encoding.hasBom) return data;
		return Buffer.concat([Buffer.from([0xfe, 0xff]), data]);
	}

	const data = Buffer.from(xml, "utf8");
	if (!encoding.hasBom) return data;
	return Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), data]);
}

export function parseXml(xml: string | Buffer): Element {
	if (Buffer.isBuffer(xml)) {
		xml = decodeXmlBuffer(xml).text;
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
	// Always use non-compact representation to match parseXml's xml2js({ compact: false })
	const convertOptions: ConvertOptions = { ...options, compact: false };
	return convertXml(parseXml(xml), convertOptions);
}

export function formatXmlBuffer(xml: Buffer, options: ConvertOptions = {}): Buffer {
	const decoded = decodeXmlBuffer(xml);
	const formatted = formatXml(decoded.text, options);
	return encodeXmlBuffer(formatted, decoded.encoding);
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
