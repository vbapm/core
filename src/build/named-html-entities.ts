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
