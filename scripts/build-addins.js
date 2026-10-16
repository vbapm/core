import { join } from "node:path";
import { fileURLToPath } from "node:url";
import * as vba from "../lib/index.js";

const __dirname = fileURLToPath(new URL(".", import.meta.url));

main().catch(err => {
	console.error(err.message);
	if (err.underlying && err.underlying.result) {
		console.log(err.underlying.result);
	}
	process.exit(1);
});

async function main() {
	// Hidden instance: a visible Excel may have the installed vbapm.xlam loaded under the same name
	await vba.buildProject({
		addin: join(__dirname, "bootstrap/build/bootstrap.xlsm"),
		background: true
	});
}
