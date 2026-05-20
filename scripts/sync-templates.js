/* eslint-disable no-console */
const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.resolve(__dirname, "..");
const TEMPLATES_DIR = path.join(ROOT, "src", "actions", "templates");
const SUBMODULE_DIR = path.join(TEMPLATES_DIR, "VBA-on-GitHub");

const TEMPLATE_MAPPINGS = [
	{
		source: path.join(SUBMODULE_DIR, "editorconfig", ".editorconfig"),
		target: path.join(TEMPLATES_DIR, "template.editorconfig")
	},
	{
		source: path.join(SUBMODULE_DIR, "gitattributes", "CRLF everywhere", ".gitattributes"),
		target: path.join(TEMPLATES_DIR, "template.gitattributes")
	},
	{
		source: path.join(SUBMODULE_DIR, "gitignore", ".gitignore"),
		target: path.join(TEMPLATES_DIR, "template.gitignore")
	}
];

function normalize(text) {
	return text.replace(/\r\n/g, "\n");
}

function ensureExists(filePath, label) {
	if (!fs.existsSync(filePath)) {
		throw new Error(`Missing ${label}: ${filePath}`);
	}
}

function readUtf8(filePath) {
	return fs.readFileSync(filePath, "utf8");
}

function writeUtf8(filePath, content) {
	fs.writeFileSync(filePath, content, "utf8");
}

function checkTemplatesSync() {
	let outOfSync = 0;

	for (const { source, target } of TEMPLATE_MAPPINGS) {
		ensureExists(source, "template source");
		ensureExists(target, "template target");

		const sourceText = normalize(readUtf8(source));
		const targetText = normalize(readUtf8(target));

		if (sourceText !== targetText) {
			outOfSync += 1;
			console.error(`Out of sync: ${path.relative(ROOT, target)}`);
		}
	}

	if (outOfSync > 0) {
		console.error("Run `npm run sync:config-templates` to update local template files.");
		process.exit(1);
	}

	console.log("Templates are in sync.");
}

function syncTemplates() {
	for (const { source, target } of TEMPLATE_MAPPINGS) {
		ensureExists(source, "template source");
		const sourceText = readUtf8(source);
		writeUtf8(target, sourceText);
		console.log(`Updated ${path.relative(ROOT, target)}`);
	}
}

function main() {
	ensureExists(SUBMODULE_DIR, "templates submodule directory");

	const checkOnly = process.argv.includes("--check");

	if (checkOnly) {
		checkTemplatesSync();
		return;
	}

	syncTemplates();
	checkTemplatesSync();
}

main();
