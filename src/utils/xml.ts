import { xml2js, js2xml, Element } from "xml-js";
import { decodeBuffer, encodeBuffer } from "./encoding";

export type Xml = any;

/**
 * Parses an XML string or buffer into an (xml-js) Element object.
 * @param xml The XML content as a string or buffer.
 * @returns The parsed XML as an (xml-js) Element object.
 */
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
	eol?: string;
}

// js2xml pre-escapes `"` → &quot; in attribute values before calling
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

/**
 * Single-pass XML formatter. Scans the input string without building a DOM
 * tree and copies all content (including entities) verbatim, so no
 * double-escaping can occur. About 10–50× faster than the parse+serialize
 * approach for large OOXML files.
 *
 * Rules:
 *  - Processing instructions and the XML declaration are emitted at depth 0.
 *  - Self-closing tags are emitted on their own line.
 *  - Elements whose only child is a text node are kept inline: <t>value</t>.
 *  - All other elements are block-formatted with increasing indentation.
 */
function formatXmlString(xml: string, indent: string, eol = "\n"): string {
	// Validate early: XML must start with a tag.
	let start = 0;
	while (start < xml.length && xml.charCodeAt(start) <= 32) start++;
	if (start >= xml.length || xml[start] !== "<") {
		throw new Error("Not valid XML: must start with <");
	}

	const N = xml.length;
	const parts: string[] = [];
	let pos = start;
	let depth = 0;

	while (pos < N) {
		// Skip inter-element whitespace (indentation / newlines from input).
		while (pos < N && xml.charCodeAt(pos) <= 32) pos++;
		if (pos >= N) break;

		if (xml[pos] !== "<") {
			// Text node: copy content and advance to next tag.
			const nextTag = xml.indexOf("<", pos);
			const text = nextTag === -1 ? xml.slice(pos) : xml.slice(pos, nextTag);
			if (text) parts.push(indent.repeat(depth), text, eol);
			pos = nextTag === -1 ? N : nextTag;
			continue;
		}

		const c1 = xml[pos + 1];

		if (c1 === "?") {
			// Processing instruction / XML declaration: <?...?>
			const end = xml.indexOf("?>", pos);
			if (end === -1) throw new Error("Unterminated processing instruction");
			parts.push(indent.repeat(depth), xml.slice(pos, end + 2), eol);
			pos = end + 2;
		} else if (c1 === "!" && xml[pos + 2] === "-" && xml[pos + 3] === "-") {
			// Comment: <!--...-->
			const end = xml.indexOf("-->", pos);
			if (end === -1) throw new Error("Unterminated comment");
			parts.push(indent.repeat(depth), xml.slice(pos, end + 3), eol);
			pos = end + 3;
		} else if (c1 === "!" && xml[pos + 2] === "[") {
			// CDATA section: <![CDATA[...]]>
			const end = xml.indexOf("]]>", pos);
			if (end === -1) throw new Error("Unterminated CDATA section");
			parts.push(indent.repeat(depth), xml.slice(pos, end + 3), eol);
			pos = end + 3;
		} else if (c1 === "/") {
			// Closing tag: </name>
			depth--;
			if (depth < 0) throw new Error("Unexpected closing tag");
			const end = xml.indexOf(">", pos);
			if (end === -1) throw new Error("Unterminated closing tag");
			parts.push(indent.repeat(depth), xml.slice(pos, end + 1), eol);
			pos = end + 1;
		} else {
			// Opening or self-closing tag.
			// Scan to '>' respecting quoted attribute values so we don't
			// mistake a '>' inside an attribute value for the tag boundary.
			let j = pos + 1;
			while (j < N) {
				if (xml[j] === ">") break;
				if (xml[j] === '"') {
					j++;
					while (j < N && xml[j] !== '"') j++;
				} else if (xml[j] === "'") {
					j++;
					while (j < N && xml[j] !== "'") j++;
				}
				j++;
			}
			if (j >= N) throw new Error("Unterminated tag");

			const tagEnd = j + 1;
			const selfClosing = xml[j - 1] === "/";

			if (selfClosing) {
				parts.push(indent.repeat(depth), xml.slice(pos, tagEnd), eol);
				pos = tagEnd;
			} else {
				// Lookahead: if the very next tag is a closing tag, this element
				// contains only text - keep it inline as <tag>text</tag>.
				const nextLt = xml.indexOf("<", tagEnd);
				if (nextLt !== -1 && xml[nextLt + 1] === "/") {
					// Preserve text exactly as-is so xml:space="preserve" content
					// (including whitespace-only strings) survives round-trips.
					const text = xml.slice(tagEnd, nextLt);
					const closeEnd = xml.indexOf(">", nextLt);
					if (closeEnd === -1) throw new Error("Unterminated closing tag");
					parts.push(
						indent.repeat(depth),
						xml.slice(pos, tagEnd),
						text,
						xml.slice(nextLt, closeEnd + 1),
						eol
					);
					pos = closeEnd + 1;
				} else {
					// Block element: opening tag on its own line then recurse deeper.
					parts.push(indent.repeat(depth), xml.slice(pos, tagEnd), eol);
					depth++;
					pos = tagEnd;
				}
			}
		}
	}

	return parts.join("");
}
/**
 * Single-pass XML formatting
 * @param xml The XML string or buffer to format.
 * @param options Formatting options.
 * @returns The formatted XML string.
 */
export function formatXml(xml: string | Buffer, options: ConvertOptions = {}): string {
	if (Buffer.isBuffer(xml)) {
		xml = decodeBuffer(xml).text;
	}
	const spaces = options.spaces ?? 2;
	const indent = typeof spaces === "string" ? spaces : " ".repeat(spaces);
	const eol = options.eol ?? "\r\n";
	return formatXmlString(xml, indent, eol);
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
