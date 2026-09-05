import walk from "walk-sync";
import { Element } from "xml-js";
import { convertXml, findElementByName, formatXmlBuffer, parseXml } from "../../utils/xml";
import { readFile, writeFile } from "../../utils/fs";
import { join } from "../../utils/path";

interface SelectionAttributes {
	activeCell?: string;
	sqref?: string;
}

interface WorksheetData {
	codeName: string;
	canonicalXml: string;
	selections: SelectionAttributes[];
}

export type PreviousWorksheetData = Map<string, WorksheetData>;

const WORKSHEET_PATH = /^xl\/worksheets\/[^/]+\.xml$/i;

function worksheetPaths(root: string): string[] {
	return walk(root, { directories: false }).filter(path => WORKSHEET_PATH.test(path));
}

function cloneWithoutSelectionAttributes(element: Element): Element {
	const clone: Element = { ...element };

	if (element.attributes) clone.attributes = { ...element.attributes };
	if (clone.name === "selection" && clone.attributes) {
		delete clone.attributes.activeCell;
		delete clone.attributes.sqref;
	}
	if (element.elements) {
		clone.elements = element.elements.map(cloneWithoutSelectionAttributes);
	}

	return clone;
}

function findElementsByName(elements: Element[] | undefined, name: string): Element[] {
	if (!elements) return [];

	return elements.flatMap(element => [
		...(element.name === name ? [element] : []),
		...findElementsByName(element.elements, name)
	]);
}

function readWorksheetData(data: Buffer): WorksheetData | undefined {
	const xml = parseXml(data);
	const worksheet = findElementByName(xml.elements, "worksheet");
	const sheetPr = findElementByName(worksheet?.elements, "sheetPr");
	const codeName = sheetPr?.attributes?.codeName;
	if (!worksheet || typeof codeName !== "string") return undefined;

	const selections = findElementsByName(worksheet.elements, "selection").map(selection => ({
		...(typeof selection.attributes?.activeCell === "string"
			? { activeCell: selection.attributes.activeCell }
			: {}),
		...(typeof selection.attributes?.sqref === "string"
			? { sqref: selection.attributes.sqref }
			: {})
	}));

	return {
		codeName,
		canonicalXml: convertXml(cloneWithoutSelectionAttributes(xml)),
		selections
	};
}

export async function readPreviousWorksheetData(root: string): Promise<PreviousWorksheetData> {
	const previous = new Map<string, WorksheetData>();

	for (const path of worksheetPaths(root)) {
		const worksheet = readWorksheetData(await readFile(join(root, path)));
		if (worksheet) previous.set(worksheet.codeName, worksheet);
	}

	return previous;
}

export async function preserveWorksheetSelections(
	root: string,
	previous: PreviousWorksheetData
): Promise<void> {
	for (const path of worksheetPaths(root)) {
		const filePath = join(root, path);
		const data = await readFile(filePath);
		const xml = parseXml(data);
		const current = readWorksheetData(data);
		const old = current && previous.get(current.codeName);

		if (!current || !old || current.canonicalXml !== old.canonicalXml) continue;

		const selections = findElementsByName(
			findElementByName(xml.elements, "worksheet")?.elements,
			"selection"
		);
		for (const [index, selection] of selections.entries()) {
			const oldSelection = old.selections[index];
			if (!selection.attributes || !oldSelection) continue;

			delete selection.attributes.activeCell;
			delete selection.attributes.sqref;
			if (oldSelection.activeCell !== undefined) {
				selection.attributes.activeCell = oldSelection.activeCell;
			}
			if (oldSelection.sqref !== undefined) {
				selection.attributes.sqref = oldSelection.sqref;
			}
		}

		await writeFile(filePath, formatXmlBuffer(Buffer.from(convertXml(xml)), {}, path));
	}
}
