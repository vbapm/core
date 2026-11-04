export function createNamedHtmlEntityReplacer(entities) {
	const entityTrie = buildEntityTrie(entities);
	return function replaceWithNamedHtmlEntities(text) {
		const characters = Array.from(text);
		const output = [];
		let replacements = 0;
		let unconverted = 0;

		for (let index = 0; index < characters.length;) {
			let node = entityTrie;
			let matchedName;
			let matchedEnd = index;

			for (let cursor = index; cursor < characters.length; cursor++) {
				node = node.children.get(characters[cursor]);
				if (!node) break;
				if (node.name) {
					matchedName = node.name;
					matchedEnd = cursor + 1;
				}
			}

			if (matchedName) {
				output.push(matchedName);
				index = matchedEnd;
				replacements++;
			} else {
				const codePoint = characters[index].codePointAt(0);
				if (codePoint > 0x7f && !(index === 0 && codePoint === 0xfeff)) unconverted++;
				output.push(characters[index]);
				index++;
			}
		}

		return { text: output.join(""), replacements, unconverted };
	};
}

export function createNamedHtmlEntityInvalidCharHandler(entities, text, onSubstitution) {
	const entityTrie = buildEntityTrie(entities);
	const characters = Array.from(text);
	const skippedIndices = new Set();

	return (char, index) => {
		if (skippedIndices.has(index)) return "";

		let node = entityTrie;
		let matchedName;
		let matchedEnd = index;

		for (let cursor = index; cursor < characters.length; cursor++) {
			node = node.children.get(characters[cursor]);
			if (!node) break;
			if (
				node.name &&
				characters.slice(index, cursor + 1).every(character => character.codePointAt(0) > 0x7f)
			) {
				matchedName = node.name;
				matchedEnd = cursor + 1;
			}
		}

		if (matchedName) {
			for (let skipped = index + 1; skipped < matchedEnd; skipped++) skippedIndices.add(skipped);
			onSubstitution?.(characters.slice(index, matchedEnd).join(""), matchedName, index);
			return matchedName;
		}

		const codePoint = char.codePointAt(0) ?? 0xfffd;
		const htmlCodePoint = codePoint >= 0xd800 && codePoint <= 0xdfff ? 0xfffd : codePoint;
		const replacement = `&#x${htmlCodePoint.toString(16)};`;
		onSubstitution?.(char, replacement, index);
		return replacement;
	};
}

function buildEntityTrie(entities) {
	const root = { children: new Map() };
	const entries = Object.entries(entities)
		.filter(([name, entity]) => name.endsWith(";") && hasNonAscii(entity.characters))
		.sort(
			([left], [right]) => left.length - right.length || (left < right ? -1 : left > right ? 1 : 0)
		);

	for (const [name, entity] of entries) {
		let node = root;
		for (const character of Array.from(entity.characters)) {
			if (!node.children.has(character)) {
				node.children.set(character, { children: new Map() });
			}
			node = node.children.get(character);
		}
		node.name ??= name;
	}

	return root;
}

function hasNonAscii(text) {
	return Array.from(text).some(character => character.codePointAt(0) > 0x7f);
}
