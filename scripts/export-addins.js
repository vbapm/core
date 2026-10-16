import { join } from "node:path";
import { fileURLToPath } from "node:url";
import * as vba from "../lib/index.js";

const __dirname = fileURLToPath(new URL(".", import.meta.url));

main().catch(err => {
	console.error(err);
	process.exit(1);
});

async function main() {
	await vba.exportProject({
		addin: join(__dirname, "bootstrap/build/bootstrap.xlsm")
	});
}
