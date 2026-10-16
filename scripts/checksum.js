import { resolve } from "node:path";
import checksum from "./lib/checksum.js";

main().catch(err => {
	console.error(err);
	process.exit(1);
});

async function main() {
	const path = process.argv[2];
	if (!path) return;

	const full_path = resolve(path);
	console.log(await checksum(full_path));
}
