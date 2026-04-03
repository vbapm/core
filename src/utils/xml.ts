import { xml2js, js2xml, Element } from "xml-js";
import { decodeBuffer, encodeBuffer } from "./encoding";

export type Xml = any;

export function parseXml(xml: string | Buffer): Element {
	if (Buffer.isBuffer(xml)) {
		xml = decodeBuffer(xml).text;
	}

	return xml2js(xml, { compact: false }) as Element;
}

export interface ConvertOptions {
	// https://github.com/nashwaan/xml-js#options-for-converting-js-object--json--xml
	compact?: boolean;
	spaces?: number | string;
}

// js2xml pre-escapes " → &quot; in attribute values before calling
// attributeValueFn, then outputs the result verbatim. This means we must NOT
// re-escape the & in those pre-added &quot; sequences. Excel XML only uses
// the four standard entities (&amp; &lt; &gt; &quot;), so the lookahead only
// needs to skip those.
function escapeXmlAttrValue(v: string): string {
	return v
		.replace(/&(?!(?:amp|lt|gt|quot);)/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;");
}

export function convertXml(value: Xml, options?: ConvertOptions): string {
	return js2xml(value, { attributeValueFn: escapeXmlAttrValue, ...options });
}

export function formatXml(xml: string | Buffer, options: ConvertOptions = {}): string {
	// Always use non-compact representation to match parseXml's xml2js({ compact: false })
	const convertOptions: ConvertOptions = { ...options, compact: false };
	return convertXml(parseXml(xml), convertOptions);
}

const MAX_FORMAT_SIZE = 5 * 1024 * 1024; // 5 MB

export function formatXmlBuffer(
	xml: Buffer,
	options: ConvertOptions = {},
	fileName?: string
): Buffer {
	if (xml.length >= MAX_FORMAT_SIZE) {
		const sizeMb = (xml.length / (1024 * 1024)).toFixed(2);
		const label = fileName ? `"${fileName}"` : "XML buffer";
		console.log(`Skipping XML formatting: ${label} size (${sizeMb} MB) exceeds the 5 MB limit.`);
		return xml;
	}
	const decoded = decodeBuffer(xml);
	const formatted = formatXml(decoded.text, options);
	return encodeBuffer(formatted, decoded.encoding);
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
