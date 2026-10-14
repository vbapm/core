const { mkdirSync, writeFileSync } = require("node:fs");
const { join, resolve } = require("node:path");

const ROOT = resolve(__dirname, "..");
const mode = process.argv[2];

if (mode !== "development" && mode !== "production") {
	console.error("Expected build mode to be development or production.");
	process.exit(1);
}

const libDirectory = join(ROOT, "lib");
mkdirSync(libDirectory, { recursive: true });
writeFileSync(join(libDirectory, ".build-mode"), `${mode}\n`, "utf8");

type CommonJsEntryPoint = never;
export type { CommonJsEntryPoint };
