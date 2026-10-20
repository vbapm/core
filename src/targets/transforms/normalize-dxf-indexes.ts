import { readdir } from "fs/promises";
import { join as nativeJoin } from "path";
import { Element } from "xml-js";
import { env } from "../../env";
import { pathExists, readFile, writeFile } from "../../utils/fs";
import { join } from "../../utils/path";
import { convertXml, findElementByName, formatXmlBuffer, parseXml } from "../../utils/xml";

const debug = env.debug("vbapm:target.transforms.normalize-dxf-indexes");
const XML_FILE = /\.(?:xml|rels)$/i;
const DXF_REFERENCE_ATTRIBUTES = new Set([
	"dxfId",
	"dataDxfId",
	"headerRowDxfId",
	"totalsRowDxfId"
]);

async function xmlPaths(root: string, directory = ""): Promise<string[]> {
	let entries;
	try {
		entries = await readdir(nativeJoin(root, directory), { withFileTypes: true });
	} catch (error) {
		if ((error as NodeJS.ErrnoException).code === "ENOENT") return [];
		throw error;
	}
	const paths = await Promise.all(
		entries.map(async entry => {
			const path = directory ? `${directory}/${entry.name}` : entry.name;
			if (entry.isDirectory()) return xmlPaths(root, path);
			return entry.isFile() && XML_FILE.test(entry.name) ? [path] : [];
		})
	);
	return paths.flat();
}

interface DxfEntry {
	index: number;
	element: Element;
	signature: string;
}

interface NormalizationResult {
	currentToNormalized: Map<number, number>;
	stylesXml: Element;
	stylesChanged: boolean;
	matched: number;
	appended: number;
	retained: number;
}

function localName(name?: string): string {
	if (!name) return "";
	const separator = name.indexOf(":");
	return separator === -1 ? name : name.slice(separator + 1);
}

function findElementDeep(elements: Element[] | undefined, name: string): Element | undefined {
	for (const element of elements ?? []) {
		if (localName(element.name) === name) return element;
		const nested = findElementDeep(element.elements, name);
		if (nested) return nested;
	}
	return undefined;
}

function namespaceDeclarations(element: Element, namespaces: Map<string, string>): void {
	for (const [name, value] of Object.entries(element.attributes ?? {})) {
		if (name === "xmlns") namespaces.set("", String(value));
		else if (name.startsWith("xmlns:")) namespaces.set(name.slice(6), String(value));
	}
}

function findInheritedNamespaces(
	elements: Element[] | undefined,
	target: Element,
	inherited = new Map<string, string>()
): Map<string, string> | undefined {
	for (const element of elements ?? []) {
		if (element === target) return inherited;
		const namespaces = new Map(inherited);
		namespaceDeclarations(element, namespaces);
		const nested = findInheritedNamespaces(element.elements, target, namespaces);
		if (nested) return nested;
	}
	return undefined;
}

function expandedName(name: string, namespaces: Map<string, string>, attribute: boolean): string {
	const separator = name.indexOf(":");
	if (separator !== -1) {
		const prefix = name.slice(0, separator);
		const namespace = namespaces.get(prefix);
		return namespace ? `{${namespace}}${name.slice(separator + 1)}` : name;
	}
	if (!attribute) {
		const namespace = namespaces.get("");
		if (namespace) return `{${namespace}}${name}`;
	}
	return name;
}

function canonicalElement(
	element: Element,
	inheritedNamespaces = new Map<string, string>()
): unknown {
	const namespaces = new Map(inheritedNamespaces);
	for (const [name, value] of Object.entries(element.attributes ?? {})) {
		if (name === "xmlns") namespaces.set("", String(value));
		else if (name.startsWith("xmlns:")) namespaces.set(name.slice(6), String(value));
	}

	const attributes = Object.entries(element.attributes ?? {})
		.filter(([name]) => name !== "xmlns" && !name.startsWith("xmlns:"))
		.map(([name, value]) => [expandedName(name, namespaces, true), value])
		.sort(([left], [right]) => String(left).localeCompare(String(right)));
	const elements = element.elements ?? [];
	const children = elements.flatMap(child => {
		if (child.type === "comment" || child.type === "instruction" || child.type === "doctype")
			return [];
		if (
			child.type === "text" &&
			elements.some(item => item.type === "element") &&
			typeof child.text === "string" &&
			!child.text.trim()
		) {
			return [];
		}
		if (child.type === "element") {
			return [["element", canonicalElement(child, namespaces)]];
		}
		return [[child.type, child.text ?? ""]];
	});

	return [expandedName(element.name ?? "", namespaces, false), attributes, children];
}

function cloneElement(element: Element): Element {
	return {
		...element,
		...(element.attributes ? { attributes: { ...element.attributes } } : {}),
		...(element.elements ? { elements: element.elements.map(cloneElement) } : {})
	};
}

function dxfEntries(stylesXml: Element): DxfEntry[] {
	const dxfs = findElementDeep(stylesXml.elements, "dxfs");
	return (dxfs?.elements ?? [])
		.filter(element => element.type === "element" && localName(element.name) === "dxf")
		.map((element, index) => ({
			index,
			element,
			signature: JSON.stringify(
				canonicalElement(element, findInheritedNamespaces(stylesXml.elements, element))
			)
		}));
}

function createNormalization(
	currentStyles: Element,
	previousStyles: Element
): NormalizationResult | undefined {
	const currentDxfs = findElementDeep(currentStyles.elements, "dxfs");
	const previousDxfs = findElementDeep(previousStyles.elements, "dxfs");
	if (!currentDxfs || !previousDxfs) return undefined;

	const currentEntries = dxfEntries(currentStyles);
	const previousEntries = dxfEntries(previousStyles);
	const previousBySignature = new Map<string, number[]>();
	for (const previous of previousEntries) {
		const matches = previousBySignature.get(previous.signature) ?? [];
		matches.push(previous.index);
		previousBySignature.set(previous.signature, matches);
	}

	const currentToNormalized = new Map<number, number>();
	const matchedPrevious = new Set<number>();
	const newEntries: DxfEntry[] = [];
	for (const current of currentEntries) {
		const matches = previousBySignature.get(current.signature);
		const previousIndex = matches?.find(index => !matchedPrevious.has(index));
		if (previousIndex === undefined) {
			newEntries.push(current);
			continue;
		}
		currentToNormalized.set(current.index, previousIndex);
		matchedPrevious.add(previousIndex);
	}

	const normalized: Array<Element | undefined> = previousEntries.map(previous => {
		const current = currentEntries.find(
			entry => currentToNormalized.get(entry.index) === previous.index
		);
		return cloneElement(current?.element ?? previous.element);
	});
	for (const current of newEntries) {
		const index = normalized.length;
		currentToNormalized.set(current.index, index);
		normalized.push(cloneElement(current.element));
	}
	const normalizedSignatures = normalized.map((element, index) => {
		const current = currentEntries.find(entry => currentToNormalized.get(entry.index) === index);
		return current?.signature ?? previousEntries[index]?.signature;
	});
	const stylesChanged =
		normalized.length !== currentEntries.length ||
		currentEntries.some((entry, index) => entry.signature !== normalizedSignatures[index]) ||
		currentDxfs.attributes?.count !== String(normalized.length);

	currentDxfs.elements = normalized.filter((element): element is Element => element !== undefined);
	currentDxfs.attributes ??= {};
	currentDxfs.attributes.count = String(normalized.length);

	return {
		currentToNormalized,
		stylesXml: currentStyles,
		stylesChanged,
		matched: matchedPrevious.size,
		appended: newEntries.length,
		retained: previousEntries.length - matchedPrevious.size
	};
}

function rewriteReferences(xml: Element, mapping: Map<number, number>, filePath: string): boolean {
	let changed = false;
	const visit = (element: Element) => {
		if (element.attributes) {
			for (const [attribute, value] of Object.entries(element.attributes)) {
				if (!DXF_REFERENCE_ATTRIBUTES.has(localName(attribute))) continue;
				if (typeof value !== "string" || !/^\d+$/.test(value)) {
					debug(`Invalid DXF reference ${attribute}=${String(value)} in ${filePath}`);
					continue;
				}
				const currentIndex = Number(value);
				const normalizedIndex = mapping.get(currentIndex);
				if (normalizedIndex === undefined) {
					debug(`Unmapped DXF reference ${attribute}=${value} in ${filePath}`);
					continue;
				}
				if (normalizedIndex !== currentIndex) {
					element.attributes[attribute] = String(normalizedIndex);
					changed = true;
				}
			}
		}
		for (const child of element.elements ?? []) visit(child);
	};
	for (const element of xml.elements ?? []) visit(element);
	return changed;
}

/** Restore DXF indexes from the previous target and remap current XML references. */
export async function normalizeDxfIndexes(
	extractedDir: string,
	previousTargetDir?: string
): Promise<void> {
	if (!previousTargetDir) return;

	const currentStylesPath = join(extractedDir, "xl", "styles.xml");
	const previousStylesPath = join(previousTargetDir, "xl", "styles.xml");
	if (!(await pathExists(currentStylesPath)) || !(await pathExists(previousStylesPath))) return;

	const currentStyles = parseXml(await readFile(currentStylesPath));
	const previousStyles = parseXml(await readFile(previousStylesPath));
	const result = createNormalization(currentStyles, previousStyles);
	if (!result) {
		debug("styles.xml has no dxfs element, skipping normalization");
		return;
	}

	const packageXmlPaths = [
		...(await xmlPaths(extractedDir, "xl/worksheets")),
		...(await xmlPaths(extractedDir, "xl/tables"))
	];
	if (result.stylesChanged) {
		const stylesOutput = formatXmlBuffer(
			Buffer.from(convertXml(result.stylesXml)),
			{},
			currentStylesPath
		);
		await writeFile(currentStylesPath, stylesOutput);
	}

	for (const path of packageXmlPaths) {
		const packagePath = path;
		if (packagePath.toLowerCase() === "xl/styles.xml") continue;
		const filePath = join(extractedDir, packagePath);
		let xml: Element;
		try {
			xml = parseXml(await readFile(filePath));
		} catch {
			debug(`Could not parse XML while checking DXF references: ${packagePath}`);
			continue;
		}
		if (!rewriteReferences(xml, result.currentToNormalized, packagePath)) continue;
		const output = formatXmlBuffer(Buffer.from(convertXml(xml)), {}, packagePath);
		await writeFile(filePath, output);
	}

	debug(
		`Normalized DXFs: ${result.matched} matched, ${result.appended} appended, ${result.retained} retained`
	);
}
