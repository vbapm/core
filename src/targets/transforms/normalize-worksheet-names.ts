/**
 * normalizeWorksheetNames
 *
 * After Excel exports a workbook it assigns worksheet XML files sequentially
 * (sheet1.xml, sheet2.xml, …) based on tab order.  When the user reorders
 * sheets the file contents are physically swapped, producing large, noisy
 * diffs even though the sheet data itself has not changed.
 *
 * This post-export step gives every worksheet file a stable, identity-based
 * name:  sht{codeName}.xml  (e.g. shtSheet1.xml, shtDashboard.xml).
 * Only the ordering metadata in workbook.xml and workbook.xml.rels needs to
 * change on a reorder — the worksheet files themselves stay untouched.
 *
 * Algorithm
 * ---------
 * 1. Read workbook.xml.rels → build rId→currentTarget map.
 * 2. For every worksheet relationship, read the worksheet XML and extract
 *    <sheetPr codeName="…"/>.  Derive the desired filename: sht{codeName}.xml.
 * 3. If the current filename already matches, skip.
 * 4. Rename the file on disk.
 * 5. Rewrite workbook.xml.rels with the updated Target values.
 * 6. Rewrite [Content_Types].xml with the updated PartName values.
 */

import { basename, join } from "../../utils/path";
import { pathExists, readFile, remove, writeFile, move } from "../../utils/fs";
import { parseXml, convertXml, findElementByName, formatXmlBuffer } from "../../utils/xml";
import { env } from "../../env";

const debug = env.debug("vbapm:target.transforms.normalize-worksheet-names");

const WORKSHEET_REL_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

interface RelEntry {
	id: string;
	type: string;
	target: string; // relative path from xl/, e.g. "worksheets/sheet1.xml"
}

/**
 * Reads <sheetPr codeName="…"/> from a worksheet XML buffer.
 * Returns undefined if not present.
 */
function extractCodeName(xmlBuf: Buffer): string | undefined {
	const xml = parseXml(xmlBuf.toString("utf8"));
	const worksheet = findElementByName(xml.elements, "worksheet");
	if (!worksheet) return undefined;
	const sheetPr = findElementByName(worksheet.elements, "sheetPr");
	if (!sheetPr || !sheetPr.attributes) return undefined;
	return sheetPr.attributes.codeName as string | undefined;
}

/**
 * Run after extraction into `extractedDir`.
 *
 * @param extractedDir  Root of the unzipped workbook (contains [Content_Types].xml,
 *                      xl/, etc.).
 */
export async function normalizeWorksheetNames(extractedDir: string): Promise<void> {
	const xlDir = join(extractedDir, "xl");
	const relsPath = join(xlDir, "_rels", "workbook.xml.rels");
	const contentTypesPath = join(extractedDir, "[Content_Types].xml");

	if (!(await pathExists(relsPath))) {
		debug("workbook.xml.rels not found, skipping normalization");
		return;
	}

	// --- 1. Parse workbook.xml.rels ---
	const relsXml = parseXml(await readFile(relsPath));
	const relationships = findElementByName(relsXml.elements, "Relationships");
	if (!relationships || !relationships.elements) {
		debug("No Relationships element found, skipping");
		return;
	}

	const rels: RelEntry[] = relationships.elements
		.filter(el => el.name === "Relationship" && el.attributes)
		.map(el => ({
			id: el.attributes!.Id as string,
			type: el.attributes!.Type as string,
			target: el.attributes!.Target as string
		}));

	const worksheetRels = rels.filter(r => r.type === WORKSHEET_REL_TYPE);
	if (worksheetRels.length === 0) {
		debug("No worksheet relationships found, skipping");
		return;
	}

	// --- 2. Determine desired filename for each worksheet ---
	// rId → desired relative target (e.g. "worksheets/shtSheet1.xml")
	const renames = new Map<string, string>(); // currentAbsPath → desiredAbsPath
	const ridToDesiredTarget = new Map<string, string>(); // rId → desired relative target

	for (const rel of worksheetRels) {
		const currentAbs = join(xlDir, rel.target);

		if (!(await pathExists(currentAbs))) {
			debug(`Worksheet file not found: ${currentAbs}, skipping`);
			ridToDesiredTarget.set(rel.id, rel.target); // keep as-is
			continue;
		}

		const buf = (await readFile(currentAbs)) as unknown as Buffer;
		const codeName = extractCodeName(buf);
		if (!codeName) {
			debug(`No codeName found in ${currentAbs}, keeping current name`);
			ridToDesiredTarget.set(rel.id, rel.target);
			continue;
		}

		const desiredFilename = `sht${codeName}.xml`;
		const existingFilename = basename(rel.target);
		const dir = rel.target.substring(0, rel.target.lastIndexOf("/") + 1); // e.g. "worksheets/"
		const desiredTarget = dir + desiredFilename;
		const desiredAbs = join(xlDir, desiredTarget);

		ridToDesiredTarget.set(rel.id, desiredTarget);

		if (existingFilename !== desiredFilename) {
			renames.set(currentAbs, desiredAbs);
			debug(`Will rename: ${existingFilename} → ${desiredFilename}`);
		}
	}

	// --- 3. Perform renames ---
	for (const [from, to] of renames) {
		if (await pathExists(to)) {
			// Destination already exists (e.g. a previous export already used the right name).
			// Remove the wrongly-named duplicate.
			await remove(from);
			debug(`Removed stale duplicate: ${from}`);
		} else {
			await move(from, to);
			debug(`Renamed: ${from} → ${to}`);
		}

		// Also rename the sidecar worksheet .rels file if present.
		// OOXML convention: xl/worksheets/_rels/<sheetFile>.rels
		const fromRels = join(from, "..", "_rels", `${basename(from)}.rels`);
		const toRels = join(to, "..", "_rels", `${basename(to)}.rels`);
		if (await pathExists(fromRels)) {
			if (await pathExists(toRels)) {
				await remove(fromRels);
				debug(`Removed stale sidecar duplicate: ${fromRels}`);
			} else {
				await move(fromRels, toRels);
				debug(`Renamed sidecar: ${basename(fromRels)} → ${basename(toRels)}`);
			}
		}
	}

	// --- 4. Rewrite workbook.xml.rels ---
	let relsChanged = false;
	for (const el of relationships.elements) {
		if (el.name !== "Relationship" || !el.attributes) continue;
		const rid = el.attributes.Id as string;
		const desired = ridToDesiredTarget.get(rid);
		if (desired && el.attributes.Target !== desired) {
			el.attributes.Target = desired;
			relsChanged = true;
		}
	}
	if (relsChanged) {
		const formatted = formatXmlBuffer(
			Buffer.from(convertXml(relsXml)),
			{ spaces: 2 },
			"workbook.xml.rels"
		);
		await writeFile(relsPath, formatted);
		debug("workbook.xml.rels updated");
	}

	// --- 5. Rewrite [Content_Types].xml ---
	if (!(await pathExists(contentTypesPath))) {
		debug("[Content_Types].xml not found, skipping content types update");
		return;
	}

	const ctXml = parseXml(await readFile(contentTypesPath));
	const types = findElementByName(ctXml.elements, "Types");
	if (!types || !types.elements) return;

	const oldParts = new Set(worksheetRels.map(r => `/xl/${r.target}`));

	let ctChanged = false;
	for (const el of types.elements) {
		if (el.name !== "Override" || !el.attributes) continue;
		const partName = el.attributes.PartName as string;
		if (!oldParts.has(partName)) continue;

		// Find the corresponding desired part name by matching filename
		const currentFilename = partName.split("/").pop()!;
		// Look up which rId had this as its original target
		const matchingRel = worksheetRels.find(r => r.target.endsWith(currentFilename));
		if (!matchingRel) continue;
		const desiredTarget = ridToDesiredTarget.get(matchingRel.id);
		if (desiredTarget === undefined) continue;
		const desired = `/xl/${desiredTarget}`;
		if (el.attributes.PartName !== desired) {
			el.attributes.PartName = desired;
			ctChanged = true;
		}
	}
	if (ctChanged) {
		const formatted = formatXmlBuffer(
			Buffer.from(convertXml(ctXml)),
			{ spaces: 2 },
			"[Content_Types].xml"
		);
		await writeFile(contentTypesPath, formatted);
		debug("[Content_Types].xml updated");
	}
}
