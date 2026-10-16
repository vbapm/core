import { mkdirSync, writeFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(fileURLToPath(new URL("..", import.meta.url)));
const mode = process.argv[2];

if (mode !== "development" && mode !== "production") {
	console.error("Expected build mode to be development or production.");
	process.exit(1);
}

const libDirectory = join(ROOT, "lib");
mkdirSync(libDirectory, { recursive: true });
writeFileSync(join(libDirectory, ".build-mode"), `${mode}\n`, "utf8");
