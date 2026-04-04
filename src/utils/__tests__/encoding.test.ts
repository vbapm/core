import { decodeBuffer, encodeBuffer } from "../encoding";

const UTF16LE_BOM = Buffer.from([0xff, 0xfe]);
const UTF16BE_BOM = Buffer.from([0xfe, 0xff]);

function toUtf16Le(text: string): Buffer {
	return Buffer.from(text, "utf16le");
}

function toUtf16Be(text: string): Buffer {
	const le = Buffer.from(text, "utf16le");
	const be = Buffer.allocUnsafe(le.length);
	for (let i = 0; i < le.length - 1; i += 2) {
		be[i] = le[i + 1];
		be[i + 1] = le[i];
	}
	return be;
}

describe("decodeBuffer", () => {
	test("strips UTF-16 LE BOM from decoded text", () => {
		const text = "hello";
		const data = Buffer.concat([UTF16LE_BOM, toUtf16Le(text)]);

		const { text: decoded, encoding } = decodeBuffer(data);

		expect(decoded).toBe(text);
		expect(encoding.encoding).toBe("utf16le");
		expect(encoding.hasBom).toBe(true);
	});

	test("strips UTF-16 BE BOM from decoded text", () => {
		const text = "hello";
		const data = Buffer.concat([UTF16BE_BOM, toUtf16Be(text)]);

		const { text: decoded, encoding } = decodeBuffer(data);

		expect(decoded).toBe(text);
		expect(encoding.encoding).toBe("utf16be");
		expect(encoding.hasBom).toBe(true);
	});

	test("decodes UTF-16 LE without BOM", () => {
		// The heuristic requires >8 null bytes at odd positions, so we need a
		// string long enough to trigger detection (>8 ASCII chars in UTF-16 LE).
		const text = "hello world";
		const data = toUtf16Le(text);

		const { text: decoded, encoding } = decodeBuffer(data);

		expect(decoded).toBe(text);
		expect(encoding.encoding).toBe("utf16le");
		expect(encoding.hasBom).toBe(false);
	});

	test("decodes UTF-8", () => {
		const text = "hello";
		const data = Buffer.from(text, "utf8");

		const { text: decoded, encoding } = decodeBuffer(data);

		expect(decoded).toBe(text);
		expect(encoding.encoding).toBe("utf8");
		expect(encoding.hasBom).toBe(false);
	});
});

describe("encodeBuffer round-trip", () => {
	test("UTF-16 LE with BOM: decode then encode produces identical buffer", () => {
		const text = "hello world";
		const original = Buffer.concat([UTF16LE_BOM, toUtf16Le(text)]);

		const { text: decoded, encoding } = decodeBuffer(original);
		const reencoded = encodeBuffer(decoded, encoding);

		expect(reencoded).toEqual(original);
	});

	test("UTF-16 BE with BOM: decode then encode produces identical buffer", () => {
		const text = "hello world";
		const original = Buffer.concat([UTF16BE_BOM, toUtf16Be(text)]);

		const { text: decoded, encoding } = decodeBuffer(original);
		const reencoded = encodeBuffer(decoded, encoding);

		expect(reencoded).toEqual(original);
	});

	test("UTF-16 LE with BOM: no double BOM in round-trip", () => {
		const text = "hello";
		const original = Buffer.concat([UTF16LE_BOM, toUtf16Le(text)]);

		const { text: decoded, encoding } = decodeBuffer(original);
		const reencoded = encodeBuffer(decoded, encoding);

		// BOM appears exactly once (bytes 0-1), not twice (bytes 0-3)
		expect(reencoded[0]).toBe(0xff);
		expect(reencoded[1]).toBe(0xfe);
		expect(reencoded[2]).toBe("h".charCodeAt(0)); // first char of text, not another BOM byte
	});
});
