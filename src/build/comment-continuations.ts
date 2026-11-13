const maxPhysicalLineLength = 1023;
const maxPhysicalLinesPerLogicalLine = 25;

export function wrapVbaComment(
	prefix: string,
	comment: string,
	lineEnding: string,
	componentName: string
): string {
	if (prefix.length + comment.length <= maxPhysicalLineLength) {
		return `${prefix}${comment}`;
	}

	const continuationBodyLength = maxPhysicalLineLength - prefix.length - 2;
	const finalBodyLength = maxPhysicalLineLength - prefix.length;
	const lines: string[] = [];
	let remainder = comment;

	while (prefix.length + remainder.length > maxPhysicalLineLength) {
		if (lines.length >= maxPhysicalLinesPerLogicalLine - 1) {
			throw new Error(
				`Comment in component "${componentName}" exceeds VBA's maximum of 24 line continuations`
			);
		}

		const [chunk, next] = takePrefix(remainder, continuationBodyLength);
		lines.push(`${prefix}${chunk} _`);
		remainder = next;
	}

	if (
		prefix.length + remainder.length > maxPhysicalLineLength ||
		remainder.length > finalBodyLength
	) {
		throw new Error(
			`Unable to wrap comment in component "${componentName}" within VBA's line limit`
		);
	}

	lines.push(`${prefix}${remainder}`);
	return lines.join(lineEnding || "\n");
}

export function unwrapVbaCommentContinuations(code: string): string {
	const lines = splitPhysicalLines(code);
	const output: string[] = [];

	for (let index = 0; index < lines.length; index++) {
		const first = lines[index];
		const prefix = getCommentPrefix(first.text);
		if (!prefix || !isWrappedContinuation(first.text)) {
			output.push(`${first.text}${first.lineEnding}`);
			continue;
		}

		let body = first.text.slice(prefix.length, -2);
		let endIndex = index;
		let continuationCount = 0;

		while (isWrappedContinuation(lines[endIndex].text)) {
			const next = lines[endIndex + 1];
			if (!next || getCommentPrefix(next.text) !== prefix) break;
			if (continuationCount >= maxPhysicalLinesPerLogicalLine - 1) {
				throw new Error("Cannot reassemble a VBA comment with more than 24 line continuations");
			}

			continuationCount++;
			endIndex++;
			body += lines[endIndex].text.slice(prefix.length);
		}

		if (endIndex === index) {
			output.push(`${first.text}${first.lineEnding}`);
			continue;
		}

		output.push(`${prefix}${body}${lines[endIndex].lineEnding}`);
		index = endIndex;
	}

	return output.join("");
}

function takePrefix(text: string, maxLength: number): [string, string] {
	let codeUnitLength = 0;
	let index = 0;

	for (const character of text) {
		if (codeUnitLength + character.length > maxLength) break;
		codeUnitLength += character.length;
		index += character.length;
	}

	return [text.slice(0, index), text.slice(index)];
}

function splitPhysicalLines(code: string): { text: string; lineEnding: string }[] {
	const lines: { text: string; lineEnding: string }[] = [];
	let start = 0;

	for (let index = 0; index < code.length; index++) {
		if (code[index] !== "\r" && code[index] !== "\n") continue;

		const lineEnding = code[index] === "\r" && code[index + 1] === "\n" ? "\r\n" : code[index];
		lines.push({ text: code.slice(start, index), lineEnding });
		index += lineEnding.length - 1;
		start = index + 1;
	}

	if (start < code.length) {
		lines.push({ text: code.slice(start), lineEnding: "" });
	}

	return lines;
}

function getCommentPrefix(line: string): string | undefined {
	return /^[ \t]*'/.exec(line)?.[0];
}

function isWrappedContinuation(line: string): boolean {
	return line.length === maxPhysicalLineLength && line.endsWith(" _");
}
