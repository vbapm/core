/**
 * Validate the integrity of a target's OOXML package before building.
 *
 * A target is an unzipped directory of OOXML parts (xl/, docProps/, etc.).
 * Before zipping it into a workbook and handing it to Excel, we verify the
 * worksheet graph is internally consistent. Excel tolerates many malformed
 * packages by "repairing" them on open, but a broken worksheet declaration can
 * surface much later as a cryptic `1004: Save method of Workbook class failed`
 * (see docs/development/veryhidden-sheet-corruption-issue.md).
 *
 * Checks performed:
 *   1. Every <sheet> in xl/workbook.xml has a non-empty r:id.
 *   2. Every r:id referenced by a <sheet> resolves to a worksheet
 *      <Relationship> in xl/_rels/workbook.xml.rels.
 *   3. Every worksheet relationship's Target points to an existing file.
 */

import dedent from "@timhall/dedent";
import { CliError, ErrorCode } from "../errors";
import { Target } from "../manifest/target";
import { pathExists, readFile } from "../utils/fs";
import { join } from "../utils/path";
import { findElementByName, parseXml } from "../utils/xml";

const WORKSHEET_REL_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

interface SheetDeclaration {
	name: string;
	relationshipId?: string;
}

/**
 * Validate the OOXML package at `targetDir` and throw a CliError with a clear
 * message if any worksheet declaration is inconsistent.
 */
export async function validateTarget(target: Target, targetDir: string): Promise<void> {
	const workbookPath = join(targetDir, "xl", "workbook.xml");
	const relsPath = join(targetDir, "xl", "_rels", "workbook.xml.rels");

	if (!(await pathExists(workbookPath))) {
		return;
	}

	const sheets = await readSheets(workbookPath);

	// 1. Every <sheet> must reference a worksheet part via a non-empty r:id.
	for (const sheet of sheets) {
		if (!sheet.relationshipId) {
			throw invalidTarget(
				target,
				dedent`
        Worksheet "${sheet.name}" in xl/workbook.xml has no relationship id (r:id).

        This is usually the result of a malformed or partially-exported target.
        Re-export the target from Excel or fix xl/workbook.xml so that every
        <sheet> entry has a non-empty r:id and a matching relationship.
      `
			);
		}
	}

	if (!(await pathExists(relsPath))) {
		throw invalidTarget(
			target,
			"xl/workbook.xml declares worksheets but xl/_rels/workbook.xml.rels is missing."
		);
	}

	// 2./3. Resolve each sheet's r:id against workbook.xml.rels and check the target file exists.
	const relationships = await readRelationships(relsPath);
	const worksheetTargets = new Map<string, string>();
	for (const rel of relationships) {
		if (rel.type === WORKSHEET_REL_TYPE) {
			worksheetTargets.set(rel.id, rel.target);
		}
	}

	const relIds = new Set<string>();
	for (const rel of relationships) {
		relIds.add(rel.id);
	}

	for (const sheet of sheets) {
		const rid = sheet.relationshipId!;

		if (!relIds.has(rid)) {
			throw invalidTarget(
				target,
				dedent`
        Worksheet "${sheet.name}" references relationship "${rid}", but no such
        relationship exists in xl/_rels/workbook.xml.rels.

        This is usually the result of a malformed or partially-exported target.
        Re-export the target from Excel or add a worksheet relationship for "${rid}".
      `
			);
		}

		const targetFile = worksheetTargets.get(rid);
		if (!targetFile) {
			throw invalidTarget(
				target,
				dedent`
        Worksheet "${sheet.name}" references relationship "${rid}", but that
        relationship is not a worksheet relationship.

        Fix xl/_rels/workbook.xml.rels so that "${rid}" points to a worksheet part.
      `
			);
		}

		const targetPath = join(targetDir, "xl", targetFile);
		if (!(await pathExists(targetPath))) {
			throw invalidTarget(
				target,
				dedent`
        Worksheet "${sheet.name}" references "${targetFile}", but that file does not
        exist in the target.

        This is usually the result of a malformed or partially-exported target.
        Re-export the target from Excel or restore the missing worksheet part.
      `
			);
		}
	}
}

interface Relationship {
	id: string;
	type: string;
	target: string;
}

async function readSheets(workbookPath: string): Promise<SheetDeclaration[]> {
	const xml = parseXml(await readFile(workbookPath));
	const workbook = findElementByName(xml.elements, "workbook");
	const sheets = findElementByName(workbook && workbook.elements, "sheets");
	if (!sheets || !sheets.elements) return [];

	const result: SheetDeclaration[] = [];
	for (const sheet of sheets.elements) {
		if (sheet.name !== "sheet") continue;
		const attrs = sheet.attributes || {};
		result.push({
			name: (attrs.name as string) ?? "(unnamed)",
			relationshipId: attrs["r:id"] as string | undefined
		});
	}

	return result;
}

async function readRelationships(relsPath: string): Promise<Relationship[]> {
	const xml = parseXml(await readFile(relsPath));
	const relationships = findElementByName(xml.elements, "Relationships");
	if (!relationships || !relationships.elements) return [];

	const result: Relationship[] = [];
	for (const rel of relationships.elements) {
		if (rel.name !== "Relationship") continue;
		const attrs = rel.attributes || {};
		result.push({
			id: (attrs.Id as string) ?? "",
			type: (attrs.Type as string) ?? "",
			target: (attrs.Target as string) ?? ""
		});
	}

	return result;
}

function invalidTarget(target: Target, message: string): CliError {
	return new CliError(
		ErrorCode.TargetInvalid,
		dedent`
      Target "${target.name}" is invalid.
      
      ${message}
    `
	);
}
