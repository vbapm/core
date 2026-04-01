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

export function convertXml(value: Xml, options?: ConvertOptions): string {
	return js2xml(value, options);
}

export function formatXml(xml: string | Buffer, options: ConvertOptions = {}): string {
	// Always use non-compact representation to match parseXml's xml2js({ compact: false })
	const convertOptions: ConvertOptions = { ...options, compact: false };
	return convertXml(parseXml(xml), convertOptions);
}

export function formatXmlBuffer(xml: Buffer, options: ConvertOptions = {}): Buffer {
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
