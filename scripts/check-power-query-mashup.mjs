import { mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";
import decompress from "decompress";
import { xml2js } from "xml-js";

function decodeXml(buffer) {
	if (buffer[0] === 0xff && buffer[1] === 0xfe) return buffer.toString("utf16le", 2);
	if (buffer[0] === 0xfe && buffer[1] === 0xff) {
		return Buffer.from(buffer.subarray(2)).swap16().toString("utf16le");
	}
	if (buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf) {
		return buffer.toString("utf8", 3);
	}
	return buffer.toString("utf8");
}

function collectElements(node, localName, matches = []) {
	if (!node || typeof node !== "object") return matches;
	if (node.type === "element" && node.name.split(":").at(-1) === localName) {
		matches.push(node);
	}
	for (const child of node.elements ?? []) collectElements(child, localName, matches);
	return matches;
}

function textContent(node) {
	if (node.type === "text" || node.type === "cdata") return node.text ?? "";
	return (node.elements ?? []).map(textContent).join("");
}

function normalizeFormula(formula) {
	return formula
		.replace(/^\uFEFF/, "")
		.replace(/\r\n?/g, "\n")
		.trim();
}

function parseSectionQueries(section) {
	const declarations = [
		...section.matchAll(/^\s*shared\s+(#"(?:[^\"]|"")*"|[A-Za-z_][\w.]*)\s*=\s*/gm)
	];
	return declarations.map((declaration, index) => {
		const identifier = declaration[1];
		const name = identifier.startsWith('#"')
			? identifier.slice(2, -1).replace(/""/g, '"')
			: identifier;
		const formulaStart = declaration.index + declaration[0].length;
		const formulaEnd = declarations[index + 1]?.index ?? section.length;
		let formula = section.slice(formulaStart, formulaEnd).trim();
		if (formula.endsWith(";")) formula = formula.slice(0, -1).trimEnd();
		return { name, formula: normalizeFormula(formula) };
	});
}

async function listFiles(directory) {
	const entries = await readdir(directory, { withFileTypes: true });
	const nested = await Promise.all(
		entries.map(entry => {
			const path = join(directory, entry.name);
			return entry.isDirectory() ? listFiles(path) : entry.isFile() ? [path] : [];
		})
	);
	return nested.flat();
}

async function readMashupQueries(xmlPath) {
	const xml = decodeXml(await readFile(xmlPath));
	const document = xml2js(xml, { compact: false });
	const mashupNodes = collectElements(document, "DataMashup");
	const queries = [];

	for (const mashupNode of mashupNodes) {
		const base64 = textContent(mashupNode).replace(/\s/g, "");
		if (!base64) continue;

		const data = Buffer.from(base64, "base64");
		if (data.length < 8) throw new Error(`Invalid DataMashup payload in ${xmlPath}`);

		const packageLength = data.readUInt32LE(4);
		const packageEnd = 8 + packageLength;
		if (packageLength === 0 || packageEnd > data.length) {
			throw new Error(`Invalid DataMashup package length in ${xmlPath}`);
		}

		const temp = await mkdtemp(join(tmpdir(), "vbapm-mashup-check-"));
		try {
			const archivePath = join(temp, "mashup.zip");
			const packageDirectory = join(temp, "package");
			await writeFile(archivePath, data.subarray(8, packageEnd));
			await decompress(archivePath, packageDirectory);

			const formulaFiles = (await listFiles(packageDirectory)).filter(path =>
				/[/\\]Formulas[/\\]Section[^/\\]*\.m$/i.test(path)
			);
			for (const formulaFile of formulaFiles) {
				const section = await readFile(formulaFile, "utf8");
				queries.push(...parseSectionQueries(section));
			}
		} finally {
			await rm(temp, { recursive: true, force: true });
		}
	}

	return queries;
}

export async function checkPowerQueryFiles(targetDirectory, sourceDirectory) {
	const customXmlDirectory = join(targetDirectory, "customXml");
	const customXmlFiles = (await listFiles(customXmlDirectory)).filter(path =>
		/[/\\]item\d+\.xml$/i.test(path)
	);
	const embeddedQueries = [];
	for (const path of customXmlFiles) embeddedQueries.push(...(await readMashupQueries(path)));

	const sourceFiles = (await listFiles(sourceDirectory)).filter(path =>
		path.toLowerCase().endsWith(".pq")
	);
	const sourceQueries = await Promise.all(
		sourceFiles.map(async path => ({
			path,
			formula: normalizeFormula(await readFile(path, "utf8"))
		}))
	);
	const matchedSources = new Set();
	const mismatches = [];

	for (const query of embeddedQueries) {
		const matchIndex = sourceQueries.findIndex(
			(source, index) => !matchedSources.has(index) && source.formula === query.formula
		);
		if (matchIndex === -1) {
			mismatches.push(`No .pq formula matches embedded query "${query.name}".`);
		} else {
			matchedSources.add(matchIndex);
		}
	}

	for (const [index, source] of sourceQueries.entries()) {
		if (!matchedSources.has(index)) {
			mismatches.push(`No embedded query matches source "${source.path}".`);
		}
	}

	return { embeddedCount: embeddedQueries.length, sourceCount: sourceQueries.length, mismatches };
}

async function main() {
	const [targetDirectory, sourceDirectory] = process.argv.slice(2);
	if (!targetDirectory || !sourceDirectory) {
		console.error(
			"Usage: node scripts/check-power-query-mashup.mjs <extracted-target-dir> <pq-source-dir>"
		);
		process.exitCode = 2;
		return;
	}

	const result = await checkPowerQueryFiles(targetDirectory, sourceDirectory);
	if (result.mismatches.length > 0) {
		for (const mismatch of result.mismatches) console.error(mismatch);
		console.error(
			`Power Query check failed: ${result.embeddedCount} embedded queries, ${result.sourceCount} source files.`
		);
		process.exitCode = 1;
		return;
	}

	console.log(`Power Query check passed: ${result.embeddedCount} formulas match.`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
	main().catch(error => {
		console.error(error instanceof Error ? error.message : String(error));
		process.exitCode = 1;
	});
}
