/**
 * Text encoding detection and buffer round-trip helpers.
 *
 * This utility was originally created to handle XML documents generated in
 * Open XML formats (e.g. OOXML parts), where files may be encoded as UTF-8,
 * UTF-16 LE, or UTF-16 BE and must keep their original encoding after edits.
 */
export type TextEncoding = "utf8" | "utf16le" | "utf16be";

export interface BufferEncoding {
	encoding: TextEncoding;
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

function detectBufferEncoding(data: Buffer): BufferEncoding {
	if (data.length >= 2) {
		if (data[0] === 0xff && data[1] === 0xfe) {
			return { encoding: "utf16le", hasBom: true };
		}

		if (data[0] === 0xfe && data[1] === 0xff) {
			return { encoding: "utf16be", hasBom: true };
		}
	}

	let oddNulls = 0;
	let evenNulls = 0;
	for (let i = 0; i < data.length; i++) {
		if (data[i] !== 0x00) continue;
		if (i % 2 === 0) {
			evenNulls++;
		} else {
			oddNulls++;
		}
	}

	// UTF-16 text without BOM often has null bytes at every other position.
	if (oddNulls > 8 && oddNulls > evenNulls * 2) {
		return { encoding: "utf16le", hasBom: false };
	}

	if (evenNulls > 8 && evenNulls > oddNulls * 2) {
		return { encoding: "utf16be", hasBom: false };
	}

	return { encoding: "utf8", hasBom: false };
}
/**
 * Decodes a UTF-8, UTF-16 LE, or UTF-16 BE buffer into a text string, along with the detected encoding and BOM presence for later re-encoding.
 * If the buffer has a BOM, it will be stripped from the output text.
 * @param data The buffer containing the encoded text.
 * @returns An object containing the decoded text and the detected encoding information.
 */
export function decodeBuffer(data: Buffer): { text: string; encoding: BufferEncoding } {
	const encoding = detectBufferEncoding(data);

	if (encoding.encoding === "utf16le") {
		const payload = encoding.hasBom ? data.subarray(2) : data;
		return { text: payload.toString("utf16le"), encoding };
	}

	if (encoding.encoding === "utf16be") {
		const payload = encoding.hasBom ? data.subarray(2) : data;
		return { text: swap16Bytes(payload).toString("utf16le"), encoding };
	}

	return { text: data.toString("utf8"), encoding };
}
/**
 * Encodes a text string into a UTF-8, UTF-16 LE, or UTF-16 BE buffer.
 * Optionally includes a BOM if specified in the encoding information.
 * @param text The text string to encode.
 * @param encoding The encoding information, including the desired encoding and BOM presence.
 * @returns A buffer containing the encoded text.
 */
export function encodeBuffer(text: string, encoding: BufferEncoding): Buffer {
	if (encoding.encoding === "utf16le") {
		const data = Buffer.from(text, "utf16le");
		if (!encoding.hasBom) return data;
		return Buffer.concat([Buffer.from([0xff, 0xfe]), data]);
	}

	if (encoding.encoding === "utf16be") {
		const data = swap16Bytes(Buffer.from(text, "utf16le"));
		if (!encoding.hasBom) return data;
		return Buffer.concat([Buffer.from([0xfe, 0xff]), data]);
	}

	const data = Buffer.from(text, "utf8");
	if (!encoding.hasBom) return data;
	return Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), data]);
}
