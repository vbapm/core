export interface NamedHtmlEntity {
	codepoints: number[];
	characters: string;
}

export interface NamedHtmlEntityReplacement {
	text: string;
	replacements: number;
	unconverted: number;
}

export type NamedHtmlEntitySubstitutionCallback = (
	original: string,
	replacement: string,
	index: number
) => void;

export function createNamedHtmlEntityReplacer(
	entities: Record<string, NamedHtmlEntity>
): (text: string) => NamedHtmlEntityReplacement;

export function createNamedHtmlEntityInvalidCharHandler(
	entities: Record<string, NamedHtmlEntity>,
	text: string,
	onSubstitution?: NamedHtmlEntitySubstitutionCallback
): (char: string, index: number) => string;
