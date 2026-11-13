import entityData from "../../scripts/data/html-entities.json" with { type: "json" };
import {
	createNamedHtmlEntityInvalidCharHandler,
	createNamedHtmlEntityReplacer,
	type NamedHtmlEntitySubstitutionCallback
} from "../../scripts/html-entities.js";

export const replaceWithNamedHtmlEntities = createNamedHtmlEntityReplacer(entityData);
export const createNamedHtmlEntityHandler = (
	text: string,
	onSubstitution?: NamedHtmlEntitySubstitutionCallback
) => createNamedHtmlEntityInvalidCharHandler(entityData, text, onSubstitution);

const namedHtmlEntities = new Map(
	Object.entries(entityData)
		.filter(
			([name, entity]) =>
				name.endsWith(";") &&
				Array.from(entity.characters).some(character => character.codePointAt(0)! > 0x7f)
		)
		.map(([name, entity]) => [name, entity.characters])
);

const htmlEntityPattern = /&(?:#(?:[xX][\da-fA-F]+|\d+)|[A-Za-z][A-Za-z\d]*);/g;
const htmlEntityWithEscapePattern =
	/&amp;(?=(?:#(?:[xX][\da-fA-F]+|\d+)|[A-Za-z][A-Za-z\d]*);)|&(?:#(?:[xX][\da-fA-F]+|\d+)|[A-Za-z][A-Za-z\d]*);/g;

export function escapeHtmlEntities(text: string): string {
	return text.replace(htmlEntityPattern, entity => `&amp;${entity.slice(1)}`);
}

export function decodeHtmlEntities(text: string): string {
	return text.replace(htmlEntityWithEscapePattern, entity => {
		if (entity === "&amp;") return "&";

		const namedEntity = namedHtmlEntities.get(entity);
		if (namedEntity !== undefined) return namedEntity;

		const numericEntity = /^&#(?:[xX]([\da-fA-F]+)|(\d+));$/.exec(entity);
		if (!numericEntity) return entity;

		const isHexadecimal = numericEntity[1] !== undefined;
		const codePoint = Number.parseInt(
			numericEntity[1] ?? numericEntity[2],
			isHexadecimal ? 16 : 10
		);
		if (!Number.isSafeInteger(codePoint) || codePoint <= 0x7f || codePoint > 0x10ffff) {
			return entity;
		}

		return codePoint >= 0xd800 && codePoint <= 0xdfff ? "\uFFFD" : String.fromCodePoint(codePoint);
	});
}
